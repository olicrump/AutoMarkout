"""
Processes a TXT file from AutoCAD's "ATTEXT" function into a markout spreadsheet.

With full GUI
"""

import csv
# Output Functions
import print as printer
import xlsxwriter
import xlsx_box
# Common / Utils / GUI
import logging
import sys
import tkinter as tk
import customtkinter
import os
import copy
import math
import traceback
from functools import partial
import subprocess, os, platform
import copy
# Get logger
logger = logging.getLogger("AutoMarkout")
# Set logging level to the logger
logger.setLevel(logging.DEBUG)

# --------------------------------------------------------------
# Classes
# --------------------------------------------------------------

class RiggingPoint:

    qut_t_names = ['250kg','250KG','0.25t','0.25T','1/4t','1/4T']
    half_t_names = ['500kg','500KG','0.5t','0.5T','1/2t','1/2T']
    one_t_names = ['1000kg','1000KG','1t','1T']
    two_t_names = ['2000kg','2000KG','2t','2T']

    def __init__(self, data=None, sorted_data=None):
        if sorted_data:
            self.layer = sorted_data['layer']
            self.block_type = sorted_data['block']
            self.point_id = sorted_data['id']
            self.point_x = sorted_data['x']
            self.point_y = sorted_data['y']
            self.point_type = sorted_data['type']
        else:
            self.layer = data[0]
            self.block_type = data[1]
            self.point_x = float(data[4])
            self.point_y = float(data[5])
            if data[2] and data[3]:
                self.point_id = data[2] + '/' + data[3]
            else:
                self.point_id = data[2] + data[3]
            if data[15] and data[16]:
                self.point_type = data[15] + '/' + data[16]
            else:
                self.point_type = data[15] + data[16]

    def print_xlsx(self,workbook, worksheet, row, col, direction, datum = (0.0,0.0)):

        if self.point_type in self.qut_t_names:
            colour = '#FDFFCC'
        elif self.point_type in self.half_t_names:
            colour = '#D8FFCC'
        elif self.point_type in self.one_t_names:
            colour = '#CCD0FF'
        elif self.point_type in self.two_t_names:
            colour = '#FFCCCC'
        else:
            colour = '#FFFFFF'

        main_format = workbook.add_format({'align': 'center',
                                           'valign': 'vcenter',
                                           'border': True,
                                           'font_size': 12,
                                             'pattern': 1,
                                           'bg_color': colour})

        meter_format = workbook.add_format({'align': 'center',
                                             'valign': 'vcenter',
                                             'border': True,
                                             'font_size': 12,
                                             'pattern': 1,
                                             'bg_color': colour,
                                             'num_format': '0.00"m"'})

        worksheet.write(row, col, self.point_id, main_format)
        worksheet.write(row+1, col, self.point_type, main_format)
        if direction == 'x':
            worksheet.write(row+2, col, (self.point_x-datum[0])/1000, meter_format)
        elif direction == 'y':
            worksheet.write(row+2, col, (self.point_y-datum[1])/1000, meter_format)
        xlsx_box.apply_outer_border_to_range(
            workbook,
            worksheet,
            {
                "first_col_index": col,
                "last_col_index": col,
                "first_row_index": row,
                "last_row_index": row+2,
                "border_style": 2,
            },
        )

class Markout:

    def __init__(self, direction, name_dict, datum=(0.0,0.0)):
        self.qtys = {'pos': 0, 'zero': 0, 'neg': 0}
        self.points = {}
        self.lines = {}
        self.name_dict = name_dict
        if direction not in ['n', 'e', 's', 'w']:
            raise ValueError('Markout direction must be one of "n", "e", "s", "w"')
        else:
            self.direction = direction

        if len(datum) != 2 and isinstance(datum[1], float) and isinstance(datum[0], float):
            raise ValueError('Markout datum must be a tuple of floats (in MM)')
        else:
            self.datum = datum

        logger.info('Init Markout Obj for [%s] with datum %s and direction %s',
                    name_dict['worksheet'], datum, direction)

    def append(self, point):

        if self.direction == 'n' or self.direction == 's':
            primary_dim = round(point.point_y,-2)
            actual_dim = point.point_y
        else:
            primary_dim = round(point.point_x,-2)
            actual_dim = point.point_x
        try:
            current_line = self.lines[primary_dim]
        except KeyError:
            current_line = {'dim':actual_dim,
                            'pos_points':[],
                            'zero_point':[],
                            'neg_points':[]}

        if self.direction == 'n' or self.direction == 's':
            if point.point_x == 0:
                current_line['zero_point'].append(point)
            elif point.point_x > 0:
                current_line['pos_points'].append(point)
            else:
                current_line['neg_points'].append(point)

            current_line['pos_points'].sort(key=lambda x: x.point_x, reverse=False)
            current_line['neg_points'].sort(key=lambda x: x.point_x, reverse=True)

            current_line_qtys = {'pos': len(current_line['pos_points']),
                                 'zero': len(current_line['zero_point']),
                                 'neg': len(current_line['neg_points'])}

            self.lines = dict(sorted(self.lines.items(), reverse=True))

        else:

            if point.point_y == 0:
                current_line['zero_point'].append(point)
            elif point.point_y > 0:
                current_line['pos_points'].append(point)
            else:
                current_line['neg_points'].append(point)

            current_line['pos_points'].sort(key=lambda x: x.point_y, reverse=False)
            current_line['neg_points'].sort(key=lambda x: x.point_y, reverse=True)

            current_line_qtys = {'pos': len(current_line['pos_points']),
                                 'zero': len(current_line['zero_point']),
                                 'neg': len(current_line['neg_points'])}

            self.lines = dict(sorted(self.lines.items()))

        for i in ['pos','zero','neg']:
            self.qtys[i] = max(current_line_qtys[i],self.qtys[i])

        self.lines[primary_dim] = current_line

        logger.info('    - Added point [%s] to markout "%s" - X=%s Y=%s',
                    ' '+point.point_id.ljust(8), self.name_dict['worksheet'], point.point_x, point.point_y)

    def print_markout(self):
        if self.points:
            dims = []
            for line in self.points:
                dims.append(line)
            for dim in dims:
                if self.direction == 'n' or self.direction == 's':
                    print('Move Tape to Y='+mm_to_m(self.points[dim][0].point_y))
                    for point in self.points[dim]:
                        print(' -- Mark point: '+point.point_id+' @ X= '+mm_to_m(point.point_x))
                else:
                    print('Move Tape to X='+mm_to_m(self.points[dim][0].point_x))
                    for point in self.points[dim]:
                        print(' -- Mark point: '+point.point_id+' @ Y= '+mm_to_m(point.point_y))

    def print_escpos(self, printer):
        if self.points:
            dims = []
            for line in self.points:
                dims.append(line)
            for dim in dims:
                if self.direction == 'n' or self.direction == 's':
                    printer.text('\nTape to  Y='+mm_to_m(self.points[dim][0].point_y)+'\n\n')
                    print('Move Tape to Y='+mm_to_m(self.points[dim][0].point_y))
                    for point in self.points[dim]:
                        printer.text(point.point_id.ljust(6)+' @ X='+mm_to_m(point.point_x).rjust(7)+'\n')
                        print(' -- Mark point: '+point.point_id+' @ X= '+mm_to_m(point.point_x))
                else:
                    printer.text('\nTape to  X='+mm_to_m(self.points[dim][0].point_x)+'\n\n')
                    print('Move Tape to X='+mm_to_m(self.points[dim][0].point_x))
                    for point in self.points[dim]:
                        printer.text(point.point_id.ljust(6)+' @ Y='+mm_to_m(point.point_y).rjust(7)+'\n')
                        print(' -- Mark point: '+point.point_id+' @ Y= '+mm_to_m(point.point_y))

    def print_xlsx(self, workbook, formats):

        if self.lines:

            logger.info('Printing Markout [%s] with lines [%s]', self.name_dict['worksheet'], self.qtys)

            worksheet = workbook.add_worksheet(self.name_dict['worksheet'])  # This adds a flysheet
            worksheet.set_paper(9)
            worksheet.center_horizontally()
            # worksheet.center_vertically()

            # middle_col = self.qtys['pos']
            # full_width = self.qtys['pos'] + self.qtys['neg']

            col = 0
            row = 3

            if self.direction == 'n' or self.direction == 's':
                worksheet.merge_range(0, 0, 0, 8, self.name_dict['worksheet'] + ' Markout', formats['city'])
                worksheet.merge_range(1, 0, 1, 8, self.name_dict['long'], formats['h1'])
                self.lines = dict(sorted(self.lines.items(), reverse=True))

                if self.datum != (0.0,0.0):
                    worksheet.merge_range(2, 0, 2, 8, 'Datum shifted to X= '+mm_to_m(self.datum[0])+' Y= '+mm_to_m(self.datum[1]), formats['h1_red'])
                    row +=1
            else:
                worksheet.merge_range(0, 0, 0, 12, self.name_dict['worksheet'] + ' Markout', formats['city'])
                worksheet.merge_range(1, 0, 1, 12, self.name_dict['long'], formats['h1'])
                self.lines = dict(sorted(self.lines.items(), reverse=False))

                if self.datum != (0.0,0.0):
                    worksheet.merge_range(2, 0, 2, 12, 'Datum shifted to X= '+mm_to_m(self.datum[0])+' Y= '+mm_to_m(self.datum[1]), formats['h1_red'])
                    row +=1


            starting_col = max(4, self.qtys['pos'])
            starting_row = 2 + (3*max(4,self.qtys['pos']))

            if self.direction == 'e' and self.name_dict['worksheet'] == 'Stage Left':
                worksheet.merge_range(starting_row, col, starting_row + 2, col, 0, formats['meter_yellow'])
                col += 1

            for dim in self.lines:
                # print(type(dim))

                if self.direction == 'n' or self.direction == 's':

                    col = starting_col

                    if self.lines[dim]['zero_point']:
                        self.lines[dim]['zero_point'][0].print_xlsx(workbook, worksheet, row, col, 'y',self.datum)
                    else:
                        worksheet.merge_range(row, col, row+2, col, (self.lines[dim]['dim']-self.datum[1])/1000,
                                              formats['meter_yellow'] if dim == 0.00 else formats['meter'])
                    col -= 1
                    for point in self.lines[dim]['neg_points']:
                        point.print_xlsx(workbook, worksheet, row, col, 'x',self.datum)
                        col -= 1
                    col = starting_col + 1
                    for point in self.lines[dim]['pos_points']:
                        point.print_xlsx(workbook, worksheet, row, col, 'x',self.datum)
                        col += 1
                    row += 3
                else:

                    # staring_row = 2 + (3*max(4,len(self.lines[dim]['pos_points'])))
                    row = starting_row

                    if self.lines[dim]['zero_point']:
                        self.lines[dim]['zero_point'][0].print_xlsx(workbook, worksheet, row, col, 'x',self.datum)
                    else:
                        worksheet.merge_range(row, col, row+2, col, (self.lines[dim]['dim']-self.datum[0])/1000,
                                              formats['meter_yellow'] if dim == 0.00 else formats['meter'])
                    row -= 3
                    for point in self.lines[dim]['pos_points']:
                        point.print_xlsx(workbook, worksheet, row, col, 'y',self.datum)
                        row -= 3
                    row = starting_row + 3
                    for point in self.lines[dim]['neg_points']:
                        point.print_xlsx(workbook, worksheet, row, col, 'y',self.datum)
                        row += 3
                    col += 1

            if self.direction == 'w' and self.name_dict['worksheet'] == 'Stage Right':
                worksheet.merge_range(starting_row, col, starting_row + 2, col, 0, formats['meter_yellow'])

class RiggingPlot:

    def __init__(self, points):
        # self.filename = filename
        # importer=ImportCSV(file)
        self.points=points
        # self.points = import_csv(filename)
        self.points.sort(key=lambda x: x.point_id, reverse=False)

        logger.info('Init RiggingPlot Obj imported %s points', len(self.points))

    def get_layers(self):
        layers = []
        for point in self.points:
            if point.layer not in layers:
                layers.append(point.layer)
        return layers

    def get_block_names(self):
        blocks = []
        for point in self.points:
            if point.block_type not in blocks:
                blocks.append(point.block_type)
        return blocks

    def get_most_us_dim(self):
        dim = None
        for point in self.points:
            if not dim:
                dim = point.point_y
            else:
                dim = max(dim, point.point_y)
        return dim

    def filter_points(self,print_states):

        self.filtered_points = []

        for point in self.points:
            if print_states[0][point.layer] == False or print_states[1][point.block_type] == False:
                pass
            else:
                self.filtered_points.append(point)

    def save_to_xlsx(self,
                     out_filename,
                     print_states=None,
                     stagesize = None,
                     delay_split = None,
                     pa_mode = False,
                     datum = (0.0,0.0)):
        if print_states:
            self.filter_points(print_states)
            markouts = assign_markouts(self.filtered_points,stagesize, delay_split, pa_mode, datum)
        else:
            markouts = assign_markouts(self.points, stagesize, delay_split, pa_mode, datum)

        workbook = xlsxwriter.Workbook(out_filename)  # This creates the Excel document
        workbook.set_size(2000, 1200)
        formats = set_formats(workbook)

        # Add sheet Properties

        workbook.set_properties({
            'title': 'AutoMarkout Data',
            'subject': 'Rigging Data',
            'author': 'AutoMarkout',
            'company': 'AutoMarkout',
            'comments': 'Created by AutoMarkout'})

        for key in markouts:
            if markouts[key]:
                markouts[key].print_xlsx(workbook, formats)

        workbook.close()  # This Shuts the document

# --------------------------------------------------------------
# Misc
# --------------------------------------------------------------

def parse_csv_to_cols(file):
    with open(file, newline='') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar="'")
        cols = []
        empty_col = {'data':[],
                     'unique_vals':[],
                     'type':None,
                     'max':None,
                     'min':None}
        for i in range(len(csvreader.__next__())):
            cols.append(copy.deepcopy(empty_col))
        for row in csvreader:
            col_no = 0
            for col in row:
                cols[col_no]['data'].append(col)
                col_no += 1

    for col in cols:
        try:
            min_val = None
            for line in col['data']:
                x=float(line.replace(' ', ''))
                if not min_val:
                    min_val = x
                    max_val = x
                else:
                    min_val = min(min_val,x)
                    max_val = max(max_val,x)
        except Exception as e:
            # print(str(type(e).__name__)+' encountered.'+ traceback.format_exc() + '\n' + str(e))
            col['type'] = 'text'
            for line in col['data']:
                if line not in col['unique_vals']:
                    col['unique_vals'].append(line)
        else:
            col['type'] = 'numeric'
            col['max'] = max_val
            col['min'] = min_val

    return cols

def import_csv(file):
    points = []
    with open(file, newline='') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar="'")
        for row in csvreader:
            if 'Datum' in row[15] or 'Datum' in row[16]:
                continue
            points.append(RiggingPoint(row))
    return points

def assign_markouts(points = [],
                    stagesize = None,
                    delay_split = None,
                    pa_mode = False,
                    datum = (0.0,0.0)):
    """
    :param points:
    A list of RiggingPoint objects
    :param stagesize:
    A tuple of floats, defining the size of the stage
    :return:
    A list of Markout objects
    0:Stage, N direction
    1:US, N direction
    2:SL, E Direction
    3:DS, S Direction
    4:SR, W Direction
    5:Delay, S Direction
    """
    markout_key = {0: {'long':  "Stage Markout, fixed tapes run US from Datum (Y-axis), Moving Tapes SL/SR (X-Axis)", # North
                       'escpos':"1: ON STAGE\nTAPES RUN US\n",
                       'worksheet':"On Stage"},
                   1: {'long':  "Upstage Markout, fixed tapes run US from Datum (Y-axis), Moving Tape US from USE (Y)", # North but should be East
                       'escpos':"2: UP STAGE\nTAPES RUN US\n",
                       'worksheet':"Upstage"},
                   2: {'long':  "Stage Left Markout, fixed tapes SL/SR (X), Moving Tapes US from DSE (Y)",
                       'escpos':"3: STAGE LEFT\nTAPES RUN SL\n",
                       'worksheet':"Stage Left"},
                   3: {'long':  "Down Stage Markout, fixed tapes US/DS (Y), Moving Tapes SL/SR (X)",
                       'escpos':"4: DOWN STAGE\nTAPES RUN DS\n",
                       'worksheet':"Downstage"},
                   4: {'long':  "Stage Right Markout, fixed tapes SL/SR (X), Moving Tapes US from DSE (Y)",
                       'escpos':"5: STAGE RIGHT\nTAPES RUN SR\n",
                       'worksheet':"Stage Right"},
                   5: {'long':  "Delays Markout, fixed tapes US/DS (Y), Moving Tapes SL/SR (X)",
                       'escpos':"6: DELAYS\nTAPES RUN DS\n",
                       'worksheet':"Delays"}}

    markouts = {}

    for point in points:

        logging.info('  - assign_markouts: Point [%s] at X=%s, Y=%s',
                     point.point_id,point.point_y, point.point_x)

        if pa_mode:
            if (point.point_y <= (- delay_split) if delay_split else False):
                if 'DL' not in markouts:
                    markouts['DL'] = Markout('s', markout_key[5], (0.0, 0.0))
                point.markout = 5
                markouts['DL'].append(point)
            else:
                if 'ST' not in markouts:
                    markouts['ST'] = Markout('e', markout_key[3],(0.0,0.0))
                # print('DS')
                point.markout = 3
                markouts['ST'].append(point)

        elif (point.point_y-datum[1]) <= 0.0:

            if (point.point_y <= (- delay_split) if delay_split else False):
                if 'DL' not in markouts:
                    markouts['DL'] = Markout('s', markout_key[5], datum)
                point.markout = 5
                markouts['DL'].append(point)
            else:
                if 'DS' not in markouts:
                    markouts['DS'] = Markout('s', markout_key[3],datum)
                # print('DS')
                point.markout = 3
                markouts['DS'].append(point)

        elif stagesize:
            if point.point_y >= stagesize[1]:
                if 'US' not in markouts:
                    markouts['US'] = Markout('e', markout_key[1],(0.0,stagesize[1]))
                # print('US')
                point.markout = 1
                markouts['US'].append(point)
            elif point.point_x >= (stagesize[0]/2):
                if 'SL' not in markouts:
                    markouts['SL'] = Markout('e', markout_key[2],(0.0,0.0))
                # print('SL')
                point.markout = 2
                markouts['SL'].append(point)
            elif point.point_x <= (-stagesize[0]/2):
                if 'SR' not in markouts:
                    markouts['SR'] = Markout('w', markout_key[4],(0.0,0.0))
                # print('SR')
                point.markout = 4
                markouts['SR'].append(point)
            else:
                if 'ST' not in markouts:
                    markouts['ST'] = Markout('n', markout_key[0], (0.0,0.0))
                # print('Stage')
                point.markout = 0
                markouts['ST'].append(point)

        else:
            if 'US' not in markouts:
                markouts['US'] = Markout('n', markout_key[1],datum)
            # print('US')
            point.markout = 1
            markouts['US'].append(point)

    return markouts

def mm_to_m(dist_mm):

    return str('{:0.2f}'.format(round(dist_mm / 1000,2))) + 'm'

def mm_to_ftin(mm):
    dec_ft = (mm/1000)*3.28084
    int_ft = int(dec_ft)
    inches = round(12*(dec_ft-int_ft),1)
    out = str(int_ft)+'\' '+str(inches)+'\"'
    return out

def set_formats(workbook):
    # Define the formats
    formats = {}
    formats['empty'] =          workbook.add_format({'border': False})

    # --------------------------------------------------------------
    # General
    # --------------------------------------------------------------

    formats['centre'] =         workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12})
    formats['centre_10'] =      workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10})
    formats['centre_10_wrap'] = workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     "text_wrap": True})
    formats['centre_big'] =     workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 14})
    formats['centre_red'] =     workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'bold': True,
                                                     'pattern': 1,
                                                     'bg_color': 'red',
                                                     'font_color': 'white'})
    formats['centre_red_10'] =  workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'bold': True,
                                                     'pattern': 1,
                                                     'bg_color': 'red',
                                                     'font_color': 'white'})
    formats['centre_yellow'] =  workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})
    formats['centre_yellow_10']=workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})
    formats['centre_grey'] =    workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'bg_color': '#D0D0D0'})

    # --------------------------------------------------------------
    # Headers
    # --------------------------------------------------------------

    formats['city'] =           workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 36,
                                                     'bold': True,
                                                     'shrink': True})

    formats['project'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 8})

    formats['centre_text'] =    workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12})
    formats['centre_text_10'] = workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10})

    # --------------------------------------------------------------
    # Panflex
    # --------------------------------------------------------------

    formats['text_01'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'bg_color': '#97d2e8'})
    formats['text_10'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'bg_color': '#97afe8'})
    formats['text_11'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'bg_color': '#9be897'})
    formats['text_12'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'bg_color': '#D3D3D3'})

    # --------------------------------------------------------------
    # Headings
    # --------------------------------------------------------------

    formats['h1'] =             workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'bold': True})
    formats['h1_10'] =          workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'bold': True})
    formats['h1_big'] =         workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 14,
                                                     'bold': True,
                                                     'underline': True})
    formats['h1_yellow'] =      workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'bold': True,
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})
    formats['h1_red'] =      workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'bold': True,
                                                     'pattern': 1,
                                                     'bg_color': 'red',
                                                     'font_color': 'white'})
    formats['h1_yellow_10'] =   workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'bold': True,
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})

    # --------------------------------------------------------------
    # Angles
    # --------------------------------------------------------------

    formats['angle'] =          workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.0"°"'})
    formats['angle_yellow'] =   workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.0"°"',
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})
    formats['angle_10'] =       workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.0"°"'})
    formats['angle_yellow_10'] =workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.0"°"',
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})

    # --------------------------------------------------------------
    # Meters
    # --------------------------------------------------------------

    formats['meter'] =          workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.00"m"'})
    formats['meter_yellow'] =   workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.00"m"',
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})
    formats['meter_10'] =       workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.00"m"'})
    formats['meter_yellow_10'] =workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.00"m"',
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})

    # --------------------------------------------------------------
    # Feet
    # --------------------------------------------------------------

    formats['foot'] =           workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.00"ft"'})
    formats['foot_yellow'] =    workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.00"ft"',
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})
    formats['foot_10'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.00"ft"'})
    formats['foot_yellow_10'] = workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.00"ft"',
                                                     'pattern': 1,
                                                     'bg_color': 'yellow'})

    # --------------------------------------------------------------
    # Kilograms
    # --------------------------------------------------------------

    formats['kg'] =             workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.0"kg"'})
    formats['kg_10'] =          workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.0"kg"'})
    formats['kg_grey'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.0"kg"',
                                                     'pattern': 1,
                                                     'bg_color': '#D0D0D0'})

    # --------------------------------------------------------------
    # Pounds
    # --------------------------------------------------------------

    formats['lb'] =             workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.0"lb"'})
    formats['lb_grey'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'num_format': '0.0"lb"',
                                                     'pattern': 1,
                                                     'bg_color': '#D0D0D0'})
    formats['lb_10'] =          workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 10,
                                                     'num_format': '0.0"lb"'})

    # --------------------------------------------------------------
    # Cable Formats
    # --------------------------------------------------------------

    formats['cable_a'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#BF8F00'})
    formats['cable_b'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#FF2600'})
    formats['cable_c'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#FF9300'})
    formats['cable_d'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#FFFC00'})
    formats['cable_e'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#00FA00'})
    formats['cable_f'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#00B0F0'})
    formats['cable_g'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#FF40FF'})
    formats['cable_h'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': '#BFBFBF'})
    formats['cable_i'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': 'white'})
    formats['cable_j'] =        workbook.add_format({'align': 'center',
                                                     'valign': 'vcenter',
                                                     'border': True,
                                                     'font_size': 12,
                                                     'pattern': 1,
                                                     'italic': True,
                                                     'bg_color': 'black',
                                                     'font_color': 'white'})

    # --------------------------------------------------------------
    # Cable List
    # --------------------------------------------------------------

    formats['cable'] = [formats['cable_a'],
                     formats['cable_b'],
                     formats['cable_c'],
                     formats['cable_d'],
                     formats['cable_e'],
                     formats['cable_f'],
                     formats['cable_g'],
                     formats['cable_h'],
                     formats['cable_i'],
                     formats['cable_j'],
                     formats['cable_a'],
                     formats['cable_b'],
                     formats['cable_c'],
                     formats['cable_d'],
                     formats['cable_e'],
                     formats['cable_f'],
                     formats['cable_g'],
                     formats['cable_h'],
                     formats['cable_i'],
                     formats['cable_j'],
                     formats['cable_a'],
                     formats['cable_b'],
                     formats['cable_c'],
                     formats['cable_d'],
                     formats['cable_e'],
                     formats['cable_f'],
                     formats['cable_g'],
                     formats['cable_h'],
                     formats['cable_i'],
                     formats['cable_j']
                     ]

    return formats

# --------------------------------------------------------------
# Print Functions
# --------------------------------------------------------------

def print_to_console(file):
    points = import_csv(file)
    points.sort(key=lambda x: x.point_id, reverse=False)
    markouts = assign_markouts(points, (18290, 9750), 24000)

    key = -1
    for markout in markouts:
        key += 1
        if markout:
            print(markout.name_dict['long'])
            markout.print_markout()

def print_to_escpos(file):
    points = import_csv(file)
    points.sort(key=lambda x: x.point_id, reverse=False)
    markouts = assign_markouts(points, (18290, 12190), 24000)

    markout_key = {0: "1: ON STAGE\nTAPES RUN US\n",
                   1: "2: UP STAGE\nTAPES RUN US\n",
                   2: "3: STAGE LEFT\nTAPES RUN SL\n",
                   3: "4: DOWN STAGE\nTAPES RUN DS\n",
                   4: "5: STAGE RIGHT\nTAPES RUN SR\n",
                   5: "6: DELAYS\nTAPES RUN DS\n"}
    with printer.escpos_printer(idVendor=0x0416, idProduct=0x5011, in_ep=130, out_ep=1) as p:

        p.printer.line_spacing(spacing=0, divisor=180)

        key = -1
        for markout in markouts:
            key += 1
            if markout:
                p.printer.set(align='center',
                                 font=0,
                                 flip=False,
                                 double_height=False,
                                 double_width=False,
                                 custom_size=True,
                                 width=2,
                                 height=2,
                                 invert=False,
                                 smooth=False,
                                 bold=False)
                p.printer.text(markout_key[key])
                print(markout_key[key])

                p.printer.set(align='center',
                                 font=1,
                                 flip=False,
                                 double_height=False,
                                 double_width=False,
                                 custom_size=True,
                                 width=2,
                                 height=1,
                                 invert=False,
                                 smooth=False,
                                 bold=False)
                markout.print_escpos(p.printer)
                p.printer.text('\n')

# --------------------------------------------------------------
# GUI
# --------------------------------------------------------------

class AutoMarkoutGUI:

    _version_num = 'v0.1.0-dev'
    _direct_print = False # True False
    _full_trace = True

    def __init__(self):

        customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
        customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

        self.ctk_btn_corner_radius = 10
        self.ctk_btn_height = 60
        self.ctk_btn_width = 60
        self.ctk_momentry_btn_fg = ('green4', 'green4')  # (light,dark)
        self.ctk_momentry_btn_hover = ('dark green', 'dark green')
        self.file_loaded = False

        # Start the window
        self.automarkout_window = customtkinter.CTk()
        self.automarkout_window.title("AutoMarkout by @oli_sound - " + self._version_num)
        self.automarkout_window.option_add("*tearOff", False)
        self.automarkout_window.lift()
        self.automarkout_window.focus_force()

        screen_width = self.automarkout_window.winfo_screenwidth()
        screen_height = self.automarkout_window.winfo_screenheight()
        x = (screen_width - (self.automarkout_window.winfo_width())) // 2
        y = (screen_height - self.automarkout_window.winfo_height()) // 3
        self.automarkout_window.geometry(f"+{x}+{y}")

        # Make the app responsive
        self.automarkout_window.columnconfigure(index=0, weight=1)
        self.automarkout_window.columnconfigure(index=1, weight=1)
        self.automarkout_window.rowconfigure(index=0, weight=1)
        self.automarkout_window.rowconfigure(index=1, weight=1)
        self.automarkout_window.rowconfigure(index=2, weight=1)
        self.automarkout_window.rowconfigure(index=3, weight=1)

        # Create control variables
        self.filename = tk.StringVar(self.automarkout_window, value='No file selected')
        self.file_selected = tk.BooleanVar(self.automarkout_window, value=False)

        self.city = tk.StringVar(self.automarkout_window, value='-')
        self.city_bits = ['-']
        self.production = tk.StringVar(self.automarkout_window, value='')
        self.units = tk.StringVar(self.automarkout_window, value="Metric (m/kg)")

        self.markout = tk.StringVar(self.automarkout_window, value="60x40ft Stage")
        self.stage_width = tk.StringVar(self.automarkout_window, value="18.29m")
        self.stage_width_flt = 18290
        self.stage_depth = tk.StringVar(self.automarkout_window, value="12.19m")
        self.stage_depth_flt = 12190

        self.stage_size = (self.stage_width_flt, self.stage_depth_flt)
        self.pa_mode = False

        self.datum = tk.StringVar(self.automarkout_window, value='Drawing Datum')
        self.datum_x = tk.StringVar(self.automarkout_window, value="0.00m")
        self.datum_x_flt = 0
        self.datum_y = tk.StringVar(self.automarkout_window, value="0.00m")
        self.datum_y_flt = 0

        # --------------------------------------------------------------
        # Create a Frame for the Import and export buttons
        # --------------------------------------------------------------
        self.import_frame = customtkinter.CTkFrame(self.automarkout_window, border_width=2, fg_color='transparent')
        self.import_frame.grid(row=0, column=0, padx=(10, 5), pady=(10, 5), sticky="nsew")

        self.import_frame.columnconfigure(index=0, weight=1, uniform="Silent_Creme")
        self.import_frame.columnconfigure(index=1, weight=1, uniform="Silent_Creme")
        self.import_frame.rowconfigure(index=0, weight=1)
        self.import_frame.rowconfigure(index=1, weight=1)
        self.import_frame.rowconfigure(index=2, weight=1)

        # Button: Open file
        self.btn_open = customtkinter.CTkButton(self.import_frame,
                                                text="Import from File",
                                                command=self.file_select,
                                                corner_radius= self.ctk_btn_corner_radius,
                                                fg_color=self.ctk_momentry_btn_fg,
                                                hover_color=self.ctk_momentry_btn_hover)
        self.btn_open.grid(row=0, column=0, columnspan=2, padx=(10,10), pady=(10,5), sticky="nsew")

        # Label: Filename
        self.lbl_file = customtkinter.CTkLabel(self.import_frame, text='Loaded file:\n'+str(self.filename.get()))
        self.lbl_file.grid(row=1, column=0, columnspan=2, padx=(10,10), pady=(5,10), sticky="nsew")

        # --------------------------------------------------------------
        # Create a Frame for the Header Text
        # --------------------------------------------------------------

        self.header_frame = customtkinter.CTkFrame(self.automarkout_window, border_width=2, fg_color='transparent')
        self.header_frame.grid(row=1, column=0, padx=(10, 5), pady=(5, 5), sticky="nsew")

        self.header_frame.columnconfigure(index=0, weight=4, uniform="Silent_Creme")
        self.header_frame.columnconfigure(index=1, weight=6, uniform="Silent_Creme")
        self.header_frame.rowconfigure(index=0, weight=1)
        self.header_frame.rowconfigure(index=1, weight=1)
        self.header_frame.rowconfigure(index=2, weight=1)

        # Label : City
        self.lbl_city = customtkinter.CTkLabel(self.header_frame,
                                               text='City Name:',
                                               justify='left')
        self.lbl_city.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky="nsew")

        # Dropdown : City
        self.dpd_city = customtkinter.CTkComboBox(self.header_frame,
                                                  values=self.city_bits,
                                                  justify='center',
                                                  variable=self.city,
                                                  state='disabled')
        self.dpd_city.grid(row=0, column=1, padx=(5,10), pady=(10,5), sticky="ew")

        # Label : Production
        self.lbl_production = customtkinter.CTkLabel(self.header_frame,
                                             text='Production:',
                                             justify='left')
        self.lbl_production.grid(row=1, column=0, padx=(10,5), pady=(5,5), sticky="nsew")

        self.txt_production = customtkinter.CTkEntry(self.header_frame,
                                             textvariable=self.production,
                                             justify="center",
                                                  state='disabled')
        self.txt_production.grid(row=1, column=1, padx=(5,10), pady=(5,10), sticky="ew")

        # --------------------------------------------------------------
        # Create a Frame for the Markout Options
        # --------------------------------------------------------------

        self.markout_frame = customtkinter.CTkFrame(self.automarkout_window, border_width=2, fg_color='transparent')
        self.markout_frame.grid(row=2, column=0, padx=(10, 5), pady=(5, 5), sticky="nsew")

        self.markout_frame.columnconfigure(index=0, weight=4, uniform="Silent_Creme")
        self.markout_frame.columnconfigure(index=1, weight=6, uniform="Silent_Creme")
        self.markout_frame.rowconfigure(index=0, weight=1)
        self.markout_frame.rowconfigure(index=1, weight=1)
        self.markout_frame.rowconfigure(index=2, weight=1)

        # Label : Filter
        self.lbl_markout = customtkinter.CTkLabel(self.markout_frame, text="Markout Type:", justify='left')
        self.lbl_markout.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky="nsew")

        self.dpd_markout = customtkinter.CTkComboBox(self.markout_frame,
                                                    values=['60x40ft Stage', '60x32ft Stage', 'Custom Stage', 'No Stage', 'PA Only Markout'],
                                                    state='disabled',
                                                    justify='center',
                                                    variable=self.markout,
                                                    command=self.markout_callback)
        self.dpd_markout.grid(row=0, column=1, padx=(5,10), pady=(10,5), sticky="ew")

        # Label : Stage Width
        self.lbl_stage_width = customtkinter.CTkLabel(self.markout_frame,
                                              text="Stage Width:",
                                              justify='right')
        self.lbl_stage_width.grid(row=1, column=0, padx=(10,5), pady=(5,5), sticky="nsew")

        # Text: Width
        self.txt_stage_width = FloatEntry(self.markout_frame,
                                          textvariable=self.stage_width,
                                                  state='disabled',
                                      text_color= 'grey')
        self.txt_stage_width.bind('<FocusOut>', self.stage_width_flush)
        self.txt_stage_width.bind('<Return>', self.stage_width_flush)
        self.txt_stage_width.grid(row=1, column=1, padx=(5,10), pady=(5,5), sticky="ew")

        # Label : Stage Depth
        self.lbl_stage_depth = customtkinter.CTkLabel(self.markout_frame,
                                              text="Stage Depth:",
                                              justify='right')
        self.lbl_stage_depth.grid(row=2, column=0, padx=(10,5), pady=(5,10), sticky="nsew")

        # Text: Width
        self.txt_stage_depth = FloatEntry(self.markout_frame,
                                          textvariable=self.stage_depth,
                                                  state='disabled',
                                      text_color= 'grey')
        self.txt_stage_depth.bind('<FocusOut>', self.stage_depth_flush)
        self.txt_stage_depth.bind('<Return>', self.stage_depth_flush)
        self.txt_stage_depth.grid(row=2, column=1, padx=(5,10), pady=(5,10), sticky="ew")

        # --------------------------------------------------------------
        # Create a Frame for the Datum options
        # --------------------------------------------------------------

        self.datum_frame = customtkinter.CTkFrame(self.automarkout_window, border_width=2, fg_color='transparent')
        self.datum_frame.grid(row=3, column=0, padx=(10, 5), pady=(5, 5), sticky="nsew")

        self.datum_frame.columnconfigure(index=0, weight=4, uniform="Silent_Creme")
        self.datum_frame.columnconfigure(index=1, weight=6, uniform="Silent_Creme")
        self.datum_frame.rowconfigure(index=0, weight=1)
        self.datum_frame.rowconfigure(index=1, weight=1)
        self.datum_frame.rowconfigure(index=2, weight=1)

        # Label : Filter
        self.lbl_datum = customtkinter.CTkLabel(self.datum_frame, text="Datum Location:", justify='left')
        self.lbl_datum.grid(row=0, column=0, padx=(10,5), pady=(10,5), sticky="nsew")

        self.dpd_datum = customtkinter.CTkComboBox(self.datum_frame,
                                                    values=['Drawing Datum', 'Most Upstage Point', 'Custom'],
                                                    state='disabled',
                                                    justify='center',
                                                    variable=self.datum,
                                                    command=self.datum_callback)
        self.dpd_datum.grid(row=0, column=1, padx=(5,10), pady=(10,5), sticky="ew")

        # Label : Stage Width
        self.lbl_datum_x = customtkinter.CTkLabel(self.datum_frame,
                                              text="Datum X (SL/SR):",
                                              justify='right')
        self.lbl_datum_x.grid(row=1, column=0, padx=(10,5), pady=(5,5), sticky="nsew")

        # Text: Width
        self.txt_datum_x = FloatEntry(self.datum_frame,
                                          textvariable=self.datum_x,
                                                  state='disabled',
                                      text_color= 'grey')
        self.txt_datum_x.bind('<FocusOut>', self.datum_x_flush)
        self.txt_datum_x.bind('<Return>', self.datum_x_flush)
        self.txt_datum_x.grid(row=1, column=1, padx=(5,10), pady=(5,5), sticky="ew")

        # Label : Stage Depth
        self.lbl_datum_y = customtkinter.CTkLabel(self.datum_frame,
                                              text="Datum Y (US/DS):",
                                              justify='right')
        self.lbl_datum_y.grid(row=2, column=0, padx=(10,5), pady=(5,10), sticky="nsew")

        # Text: Width
        self.txt_datum_y = FloatEntry(self.datum_frame,
                                      textvariable=self.datum_y,
                                      state='disabled',
                                      text_color= 'grey')
        self.txt_datum_y.bind('<FocusOut>', self.datum_y_flush)
        self.txt_datum_y.bind('<Return>', self.datum_y_flush)
        self.txt_datum_y.grid(row=2, column=1, padx=(5,10), pady=(5,10), sticky="ew")

        # --------------------------------------------------------------
        # Create a Frame for Block Selection
        # --------------------------------------------------------------

        self.selection_frame = customtkinter.CTkFrame(self.automarkout_window, fg_color='transparent')#, border_width=2)
        self.selection_frame.grid(row=0, column=1, rowspan=5, padx=(5, 10), pady=(10, 10), sticky="nsew")

        self.block_frame_out = customtkinter.CTkFrame(self.selection_frame, border_width=2, fg_color='transparent')
        self.block_frame_out.grid(row=0, column=0, padx=(0, 0), pady=(0, 5), sticky="nsew")

        # Label: Filename
        self.lbl_block = customtkinter.CTkLabel(self.block_frame_out, text='Block Selection:')
        self.lbl_block.grid(row=0, column=0, padx=(10,10), pady=(10,5), sticky="nsew")

        self.block_frame = customtkinter.CTkFrame(self.block_frame_out, border_width=2, fg_color='transparent')
        self.block_frame.grid(row=1, column=0, padx=(10, 10), pady=(5, 10), sticky="nsew")

        self.lst_block_selection = ScrollableCheckBoxFrame(master=self.block_frame,
                                                                 width = 450,
                                                                 #height= 250,
                                                                 item_list=[])

        # --------------------------------------------------------------
        # Create a Frame for Layer Selection
        # --------------------------------------------------------------

        self.layer_frame_out = customtkinter.CTkFrame(self.selection_frame, border_width=2, fg_color='transparent')
        self.layer_frame_out.grid(row=1, column=0, padx=(0, 0), pady=(5, 0), sticky="nsew")

        # Label: Filename
        self.lbl_layer = customtkinter.CTkLabel(self.layer_frame_out, text='Layer Selection:')
        self.lbl_layer.grid(row=2, column=0, padx=(10,10), pady=(10,5), sticky="nsew")

        self.layer_frame = customtkinter.CTkFrame(self.layer_frame_out, border_width=2, fg_color='transparent')
        self.layer_frame.grid(row=3, column=0, padx=(10, 10), pady=(5, 10), sticky="nsew")

        self.lst_layer_selection = ScrollableCheckBoxFrame(master=self.layer_frame,
                                                                 width = 450,
                                                                 #height= 250,
                                                                 item_list=[])

        # --------------------------------------------------------------
        # Create a Frame for Export
        # --------------------------------------------------------------

        self.export_frame = customtkinter.CTkFrame(self.automarkout_window, border_width=2, fg_color='transparent')
        self.export_frame.grid(row=4, column=0, padx=(10, 5), pady=(5, 10), sticky="nsew")

        self.export_frame.columnconfigure(index=0, weight=4, uniform="Silent_Creme")
        self.export_frame.columnconfigure(index=1, weight=6, uniform="Silent_Creme")
        self.export_frame.rowconfigure(index=0, weight=1)
        self.export_frame.rowconfigure(index=1, weight=1)

        # Label : Unit Selection
        self.lbl_units = customtkinter.CTkLabel(self.export_frame,
                                                text="Export Units:",
                                                justify='left')
        self.lbl_units.grid(row=0, column=0, padx=(10,5), pady=(10, 5), sticky="nsew")

        self.dpd_units = customtkinter.CTkComboBox(self.export_frame,
                                                   values=['Metric (m/kg)', 'Imperial (ft/lb)', 'Both Metric & Imperial'],
                                                   state='disabled',
                                                   justify='center',
                                                   variable=self.units)
        self.dpd_units.grid(row=0, column=1, padx=(5,10), pady=(10, 5), sticky="ew")

        if self._direct_print == True:

            # Button: Export to XLSX
            self.btn_export = customtkinter.CTkButton(self.export_frame,
                                                      text="Export to Excel",
                                                      command=self.export_xlsx,
                                                      corner_radius=self.ctk_btn_corner_radius,
                                                      fg_color=self.ctk_momentry_btn_fg,
                                                      hover_color=self.ctk_momentry_btn_hover,
                                                      state='disabled')
            self.btn_export.grid(row=1, column=0, padx=(10, 5), pady=(5, 10), sticky="nsew")

            # Button: Print
            self.btn_print = customtkinter.CTkButton(self.export_frame,
                                                      text="Print to Reciept",
                                                      command=self.printer_gui_try,
                                                      corner_radius=self.ctk_btn_corner_radius,
                                                      fg_color=self.ctk_momentry_btn_fg,
                                                      hover_color=self.ctk_momentry_btn_hover,
                                                      state='disabled')
            self.btn_print.grid(row=1, column=1, padx=(5, 10), pady=(5, 10), sticky="nsew")

        else:

            # Button: Export to XLSX
            self.btn_export = customtkinter.CTkButton(self.export_frame,
                                                      text="Export selected points to Excel",
                                                      command=self.export_xlsx,
                                                      corner_radius= self.ctk_btn_corner_radius,
                                                      fg_color=self.ctk_momentry_btn_fg,
                                                      hover_color=self.ctk_momentry_btn_hover,
                                                      state='disabled')
            self.btn_export.grid(row=1, column=0, columnspan=2, padx=(10,10), pady=(5,10), sticky="nsew")

        # --------------------------------------------------------------
        # Admin / Mainloop
        # --------------------------------------------------------------

        screen_width = self.automarkout_window.winfo_screenwidth()
        screen_height = self.automarkout_window.winfo_screenheight()
        x = (screen_width - (self.automarkout_window.winfo_width())) // 2
        y = (screen_height - self.automarkout_window.winfo_height()) // 3
        self.automarkout_window.geometry(f"+{x}+{y}")

        self.automarkout_window.mainloop()

    def setwritestate(self, filter):

        if filter == 'All Sources':
            filter = 0
        elif filter == 'Only Flown':
            filter = 1
        elif filter == 'Only Flown & X-Pos':
            filter = 2
        elif filter == 'None':
            filter = 3
        else:
            return

        for source in self.data['sources']:
            if filter == 0:
                source['write'] = True
            elif filter == 3:
                source['write'] = False
            elif 'flown' not in source['config']:
                source['write'] = False
            elif filter == 1:
                source['write'] = True
            elif source['position']['x'] >= 0:
                source['write'] = True
            else:
                source['write'] = False

    # --------------------------------------------------------------
    # Load button Functions
    # --------------------------------------------------------------

    def file_select(self):
        logging.debug('  - Gui: File Selection Opened')
        filetypes = (('ATTEXT CDF Files', '*.TXT'),
                     ('All files', '*.*'))
        self.filename = tk.filedialog.askopenfilename(title='Select an Attribute Export File',
                                           initialdir='/',
                                           filetypes=filetypes)
        if len(self.filename) != 0:
            self.load_file()

    def load_file(self):

        # Check what type of file and export the data:

        if '.txt' in self.filename:
            logging.info(' -- Gui: Filename Returned: %s', self.filename)
            try:

                importer=ImportCSV(self.filename, self.automarkout_window)
                points=importer.return_points()
                self.rigging_plot = RiggingPlot(points)
            except Exception as e:
                tk.messagebox.showerror(title='Error',
                                        message=str(type(e).__name__)+' encountered while Processing File',
                                        detail=e if self._full_trace == False else traceback.format_exc() + '\n' + str(e))
            else:
                self.load_data()

        else:
            logging.info(' -- Not TXT Error')
            tk.messagebox.showerror(title='Error',
                                    message='This is not a valid file',
                                    detail='Please select a .txt file')
            return None

    def load_data(self):

        # Start sorting the data
        self.city_bits = os.path.split(self.filename)[1].replace('_', ' ').replace('.', ' ').replace('-', ' ')
        self.city_bits = self.city_bits.split(' ')
        # for bit in self.city_bits:
        #     if len(bit) < 2:
        #         self.city_bits.remove(bit)
        try:
            self.city_bits.remove('txt')
        except ValueError:
            pass

        if len(self.city_bits) > 1:
            self.city = (self.city_bits[1])
        else:
            self.city = (self.city_bits[0])

        # Update fields in GUI

        self.lbl_file.configure(text= 'Loaded file:\n'+self.filename.split('/')[-1])

        self.dpd_city.configure(variable = tk.StringVar(self.automarkout_window,
                                                        value=self.city),
                                values = self.city_bits)

        # self.update_checkboxes(self.filter.get())

        self.lst_block_selection.clear_all()

        self.lst_layer_selection.clear_all()

        for block in self.rigging_plot.get_block_names():
            # Add the source to the list, and pre-set the checkbox
            self.lst_block_selection.add_item(block,True)

        for layer in self.rigging_plot.get_layers():
            # Add the source to the list, and pre-set the checkbox
            self.lst_layer_selection.add_item(layer,True)

        # self.dpd_city.configure(state='normal')
        # self.txt_production.configure(state='normal')
        # self.dpd_units.configure(state='readonly')
        self.dpd_markout.configure(state='readonly')
        # self.dpd_datum.configure(state='readonly')

        self.btn_export.configure(state='normal')
        if self._direct_print == True:
            self.btn_print.configure(state='normal')

    # --------------------------------------------------------------
    # Export button Functions
    # --------------------------------------------------------------

    def export_xlsx(self):
        self.out_filename = tk.filedialog.asksaveasfilename(title='Save As',
                                            initialdir=os.path.split(self.filename)[0],
                                            initialfile=os.path.split(self.filename)[1].replace('.txt','_AutoMarkout'),
                                            defaultextension=".xlsx")

        self.city = self.dpd_city.get()
        self.production = self.txt_production.get()

        if self.dpd_units.get() == 'Metric (m/kg)':
            units = 0
        elif self.dpd_units.get() == 'Imperial (ft/lb)':
            units = 1
        else:
            units = 2

        try:
            self.rigging_plot.save_to_xlsx(self.out_filename,
                                           print_states=self.read_checkboxes(),
                                           stagesize=self.stage_size,
                                           delay_split=24000 if self.pa_mode else None,
                                           pa_mode=self.pa_mode,
                                           datum=(self.datum_x_flt,self.datum_y_flt))
        except Exception as e:
            tk.messagebox.showerror(title='Error',
                                    message=str(type(e).__name__)+' encountered while Saving Excel File',
                                    detail=e if self._full_trace == False else traceback.format_exc() + '\n' + str(e))
        else:
            answer = tk.messagebox.askquestion(title='Finished',
                                      message='The selected sources have been exported successfully',
                                      detail='Would you like to open the spreadsheet?',
                                      type='yesno')
            if answer == 'yes':
                self.open_spreadsheet()

    def open_spreadsheet(self):
        import subprocess, os, platform
        if platform.system() == 'Darwin':  # macOS
            subprocess.call(('open', self.out_filename))
        elif platform.system() == 'Windows':  # Windows
            os.startfile(os.path.abspath(self.out_filename))
        else:  # linux variants
            subprocess.call(('xdg-open', self.out_filename))

    def read_checkboxes(self):

        self.layer_print_state = {}
        self.block_print_state = {}

        for layer in self.rigging_plot.get_layers():
            self.layer_print_state[layer] = self.lst_layer_selection.get(layer).get()
        for block in self.rigging_plot.get_block_names():
            self.block_print_state[block] = self.lst_block_selection.get(block).get()

        return [self.layer_print_state,self.block_print_state]

    # --------------------------------------------------------------
    # Variable entry Validation
    # --------------------------------------------------------------

    def stage_width_flush(self, event):
        entry = self.stage_width.get().replace('m', '')
        try:
            self.stage_width_flt = int(1000*float(entry))
        except:
            raise TypeError
        else:
            self.stage_width.set(str(f'{(self.stage_width_flt/1000):.2f}') + "m")

        self.stage_size = (self.stage_width_flt, self.stage_depth_flt)

    def stage_depth_flush(self, event):
        entry = self.stage_depth.get().replace('m', '')
        try:
            self.stage_depth_flt = int(1000*float(entry))
        except:
            raise TypeError
        else:
            self.stage_depth.set(str(f'{(self.stage_depth_flt/1000):.2f}') + "m")

        self.stage_size = (self.stage_width_flt, self.stage_depth_flt)

    def datum_x_flush(self, event):
        entry = self.datum_x.get().replace('m', '')
        try:
            self.datum_x_flt = int(1000*float(entry))
        except:
            raise TypeError
        else:
            self.datum_x.set(str(f'{(self.datum_x_flt/1000):.2f}') + "m")

    def datum_y_flush(self, event):
        entry = self.datum_y.get().replace('m', '')
        try:
            self.datum_y_flt = int(1000*float(entry))
        except:
            raise TypeError
        else:
            self.datum_y.set(str(f'{(self.datum_y_flt/1000):.2f}') + "m")

    # --------------------------------------------------------------
    # Dropdown Callbacks
    # --------------------------------------------------------------

    def markout_callback(self, event):
        markout_type = self.dpd_markout.get()
        if markout_type == '60x40ft Stage':
            self.txt_stage_width.configure(state='disabled', text_color= 'grey')
            self.txt_stage_depth.configure(state='disabled', text_color= 'grey')

            self.dpd_datum.configure(state='disabled')
            self.datum_x_flt = 0
            self.datum_x.set(str(f'{(self.datum_x_flt/1000):.2f}') + "m")
            self.datum_y_flt = 0
            self.datum_y.set(str(f'{(self.datum_y_flt/1000):.2f}') + "m")

            self.stage_width_flt = 18290
            self.stage_depth_flt = 12190

            self.stage_width.set(str(f'{(self.stage_width_flt/1000):.2f}') + "m")
            self.stage_depth.set(str(f'{(self.stage_depth_flt/1000):.2f}') + "m")

            self.stage_size = (self.stage_width_flt, self.stage_depth_flt)
            self.pa_mode = False

        elif markout_type == '60x32ft Stage':
            self.txt_stage_width.configure(state='disabled', text_color= 'grey')
            self.txt_stage_depth.configure(state='disabled', text_color= 'grey')

            self.dpd_datum.configure(state='disabled')
            self.datum_x_flt = 0
            self.datum_x.set(str(f'{(self.datum_x_flt/1000):.2f}') + "m")
            self.datum_y_flt = 0
            self.datum_y.set(str(f'{(self.datum_y_flt/1000):.2f}') + "m")

            self.stage_width_flt = 18290
            self.stage_depth_flt = 9750

            self.stage_width.set(str(f'{(self.stage_width_flt/1000):.2f}') + "m")
            self.stage_depth.set(str(f'{(self.stage_depth_flt/1000):.2f}') + "m")

            self.stage_size = (self.stage_width_flt, self.stage_depth_flt)
            self.pa_mode = False

        elif markout_type == 'Custom Stage':
            self.txt_stage_width.configure(state='normal', text_color= 'white')
            self.txt_stage_depth.configure(state='normal', text_color= 'white')

            self.dpd_datum.configure(state='disabled')
            self.datum_x_flt = 0
            self.datum_x.set(str(f'{(self.datum_x_flt/1000):.2f}') + "m")
            self.datum_y_flt = 0
            self.datum_y.set(str(f'{(self.datum_y_flt/1000):.2f}') + "m")

            self.pa_mode = False

        elif markout_type == 'No Stage':
            self.txt_stage_width.configure(state='disabled', text_color= 'grey')
            self.txt_stage_depth.configure(state='disabled', text_color= 'grey')
            self.dpd_datum.configure(state='readonly')

            self.stage_width_flt = 0
            self.stage_depth_flt = 0

            self.stage_width.set(str(f'{(self.stage_width_flt/1000):.2f}') + "m")
            self.stage_depth.set(str(f'{(self.stage_depth_flt/1000):.2f}') + "m")

            self.stage_size = None
            self.pa_mode = False

        elif markout_type == 'PA Only Markout':
            self.txt_stage_width.configure(state='disabled', text_color= 'grey')
            self.txt_stage_depth.configure(state='disabled', text_color= 'grey')
            self.dpd_datum.configure(state='disabled')

            self.stage_width_flt = 0
            self.stage_depth_flt = 0

            self.stage_width.set(str(f'{(self.stage_width_flt/1000):.2f}') + "m")
            self.stage_depth.set(str(f'{(self.stage_depth_flt/1000):.2f}') + "m")

            self.stage_size = None
            self.pa_mode = True

    def datum_callback(self, event):
        datum_type = self.dpd_datum.get()
        if datum_type == 'Drawing Datum':
            self.txt_datum_x.configure(state='disabled', text_color= 'grey')
            self.txt_datum_y.configure(state='disabled', text_color= 'grey')

            self.datum_x_flt = 0
            self.datum_y_flt = 0

            self.datum_y.set(str(f'{(self.datum_x_flt/1000):.2f}') + "m")
            self.datum_y.set(str(f'{(self.datum_y_flt/1000):.2f}') + "m")

        elif datum_type == 'Most Upstage Point':
            self.txt_datum_x.configure(state='disabled', text_color= 'grey')
            self.txt_datum_y.configure(state='disabled', text_color= 'grey')

            self.datum_x_flt = 0
            self.datum_y_flt = self.rigging_plot.get_most_us_dim()

            self.datum_y.set(str(f'{(self.datum_x_flt/1000):.2f}') + "m")
            self.datum_y.set(str(f'{(self.datum_y_flt/1000):.2f}') + "m")

        elif datum_type == 'Custom':
            self.txt_datum_x.configure(state='normal', text_color= 'white')
            self.txt_datum_y.configure(state='normal', text_color= 'white')

    # --------------------------------------------------------------
    # Print button Functions
    # --------------------------------------------------------------

    def printer_gui_try(self):

        try:
            self.printer_gui()
        except Exception as e:
            tk.messagebox.showerror(title='Error',
                                    message=str(type(e).__name__)+' encountered.',
                                    detail=e if self._full_trace == False else traceback.format_exc() + '\n' + str(e))

    def printer_gui(self):

        # Start the window
        self.print_window = customtkinter.CTk()
        self.print_window.title("Direct Print")
        self.print_window.option_add("*tearOff", False)
        self.print_window.lift()
        self.print_window.focus_force()

        screen_width = self.print_window.winfo_screenwidth()
        screen_height = self.print_window.winfo_screenheight()
        x = (screen_width - (self.print_window.winfo_width() + 300)) // 2
        y = (screen_height - self.print_window.winfo_height()) // 2
        self.print_window.geometry(f"+{x}+{y}")

        # Make the app responsive
        self.print_window.columnconfigure(index=0, weight=1)
        self.print_window.columnconfigure(index=1, weight=1)
        self.print_window.rowconfigure(index=0, weight=1)
        self.print_window.rowconfigure(index=1, weight=1)
        self.print_window.rowconfigure(index=2, weight=1)
        self.print_window.rowconfigure(index=3, weight=1)

        # Create control variables
        self.tk_idVendor = tk.StringVar(self.print_window, value="0x0485")
        self.tk_idProduct = tk.StringVar(self.print_window, value="0x7541")
        self.tk_in_ep = tk.StringVar(self.print_window, value="130")
        self.tk_out_ep = tk.StringVar(self.print_window, value="1")
        self.tk_device = tk.StringVar(self.automarkout_window, value="-")

        self.devices = printer.device_finder()
        self.device_names = ["-"]
        for device in self.devices:
            self.device_names.append(device["product"])

        # --------------------------------------------------------------
        # Header
        # --------------------------------------------------------------

        self.lbl_head = customtkinter.CTkLabel(self.print_window,
                                                 text='This allows you to print fly data directly to supported\n'
                                                      'thermal reciept printers. To start, Select your printer\n'
                                                      'from the device list and press "Print" to send the data.',
                                                 justify='left')
        self.lbl_head.grid(row=0, column=0, padx=(10,5), pady=(5,5), sticky="nsew", columnspan=2)

        # self.print_window.columnconfigure(index=0, weight=6, uniform="Silent_Creme")
        # self.print_window.columnconfigure(index=1, weight=6, uniform="Silent_Creme")

        self.lbl_device = customtkinter.CTkLabel(self.print_window,
                                                 text='Device:',
                                                 justify='left')
        self.lbl_device.grid(row=1, column=0, padx=(10,5), pady=(5,5), sticky="nsew")

        self.dpd_device = customtkinter.CTkComboBox(self.print_window,
                                                    values=self.device_names,
                                                    variable=self.tk_device,
                                                    state="readonly")
        self.dpd_device.grid(row=1, column=1, padx=(5,10), pady=(5,5), sticky="ew")

        # --------------------------------------------------------------
        # Print Buttons
        # --------------------------------------------------------------

        self.outer_print_frame = customtkinter.CTkFrame(self.print_window, border_width=2, fg_color='transparent')
        self.outer_print_frame.grid(row=2, column=0, padx=(10, 10), pady=(5, 5), sticky="nsew", columnspan=2)

        self.print_frame = customtkinter.CTkScrollableFrame(self.outer_print_frame,
                                                                 width = 300,
                                                                 height= 250,)
        self.print_frame.pack(expand=True, fill='both') #.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")


        self.print_frame.columnconfigure(index=0, weight=6, uniform="Silent_Creme")
        self.print_frame.columnconfigure(index=1, weight=6, uniform="Silent_Creme")

        self.read_checkboxes()
        frame_row = 0

        source_numbers = list(range(len(self.data['sources'])))
        for current_no in source_numbers:
            if self.data['sources'][current_no]['write']==True:

                # Label: Source Name
                self.lbl_file = customtkinter.CTkLabel(self.print_frame,
                                                       text=str(self.data['sources'][current_no]['name']))
                self.lbl_file.grid(row=frame_row, column=0, padx=(10,5), pady=(5, 5), sticky="nsew")

                # Button: Open file
                self.btn_open = customtkinter.CTkButton(self.print_frame,
                                                        text="Print",
                                                        command=partial(self.print_escpos, current_no),
                                                        corner_radius=self.ctk_btn_corner_radius,
                                                        fg_color=self.ctk_momentry_btn_fg,
                                                        hover_color=self.ctk_momentry_btn_hover)
                self.btn_open.grid(row=frame_row, column=1, padx=(5, 10), pady=(5, 5), sticky="nsew")

                frame_row += 1

        # --------------------------------------------------------------
        # Settings
        # --------------------------------------------------------------
        self.settings_frame = customtkinter.CTkFrame(self.print_window, border_width=2, fg_color='transparent')
        self.settings_frame.grid(row=3, column=0, padx=(10, 10), pady=(5, 10), sticky="nsew", columnspan=2)

        self.settings_frame.columnconfigure(index=0, weight=6, uniform="Silent_Creme")
        self.settings_frame.columnconfigure(index=1, weight=6, uniform="Silent_Creme")
        self.settings_frame.rowconfigure(index=0, weight=1)
        self.settings_frame.rowconfigure(index=1, weight=1)
        self.settings_frame.rowconfigure(index=3, weight=1)

        # Label : Warning
        self.lbl_filter = customtkinter.CTkLabel(self.settings_frame,
                                                 text="Advanced Settings:",
                                                 justify='center', text_color='grey')
        self.lbl_filter.grid(row=0, column=0, padx=(10,10), pady=5, sticky="nsew", columnspan=2)

        # Label : in_ep
        self.lbl_in_ep = customtkinter.CTkLabel(self.settings_frame,
                                                  text='Input Endpoint:',
                                                  justify='left')
        self.lbl_in_ep.grid(row=1, column=0, padx=(10,5), pady=5, sticky="nsew")

        self.txt_in_ep = customtkinter.CTkEntry(self.settings_frame,
                                                  textvariable=self.tk_in_ep,
                                                  justify="center")
        self.txt_in_ep.grid(row=1, column=1, padx=(5,10), pady=5, sticky="nsew")

        # Label : out_ep
        self.lbl_out_ep = customtkinter.CTkLabel(self.settings_frame,
                                                  text='Output Endpoint:',
                                                  justify='left')
        self.lbl_out_ep.grid(row=2, column=0, padx=(10,5), pady=(5,10), sticky="nsew")

        self.txt_out_ep = customtkinter.CTkEntry(self.settings_frame,
                                                  textvariable=self.tk_out_ep,
                                                  justify="center")
        self.txt_out_ep.grid(row=2, column=1, padx=(5,10), pady=(5,10), sticky="nsew")

        # --------------------------------------------------------------
        # Admin/Mainloop
        # --------------------------------------------------------------

        screen_width = self.print_window.winfo_screenwidth()
        screen_height = self.print_window.winfo_screenheight()
        x = (screen_width - (self.print_window.winfo_width()-300)) // 2
        y = (screen_height - self.print_window.winfo_height()-200) // 2
        self.print_window.geometry(f"+{x}+{y}")

        self.print_window.mainloop()

    def print_escpos(self, array):

        self.chosen_device = self.dpd_device.get()
        if self.chosen_device == '-' or self.chosen_device == '':
            tk.messagebox.showerror(title='Error',
                                    message='No Printer Selected',
                                    detail="Please select a printer before trying to print")
        else:
            for device in self.devices:
                if str(device['product']) == str(self.chosen_device):
                    self.idVendor=device["idVendor"]
                    self.idProduct=device["idProduct"]

            self.in_ep = int(self.tk_in_ep.get())
            self.out_ep = int(self.tk_out_ep.get())

            self.data['project']['city'] = self.dpd_city.get()

            try:
                # with escpos_printer(idVendor=0x0485, idProduct=0x7541, in_ep=130, out_ep=3) as printer:
                with escpos_printer(idVendor=self.idVendor, idProduct=self.idProduct, in_ep=self.in_ep, out_ep=self.out_ep) as printer:
                    printer.print_source(self.data,array)
            except Exception as e:
                tk.messagebox.showerror(title='Error',
                                        message=str(type(e).__name__)+' encountered while Printing',
                                        detail=e if self._full_trace == False else traceback.format_exc() + '\n' + str(e))

class ImportCSV():

    def __init__(self, filename, parent):

        self.parent = parent
        self.filename = filename

        logger.info('Importer Initialization for file: "%s"', filename)

        customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
        customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

        self.ctk_btn_corner_radius = 10
        self.ctk_btn_height = 60
        self.ctk_btn_width = 60
        self.ctk_momentry_btn_fg = ('green4', 'green4')  # (light,dark)
        self.ctk_momentry_btn_hover = ('dark green', 'dark green')

        self.block_selected = False
        self.layer_selected = False
        self.pntid_selected = False
        self.ptype_selected = False
        self.pontx_selected = False
        self.ponty_selected = False

        self.block_column_no = None
        self.layer_column_no = None

        # self.data = parse_csv_to_cols(filename)
        # self.block_finder_gui()

    def block_finder_gui(self):

        # Start the window
        # self.blockfinder_window = customtkinter.CTk()
        self.blockfinder_window = customtkinter.CTkToplevel(self.parent)
        self.blockfinder_window.title("Import CSV Step 1")
        self.blockfinder_window.option_add("*tearOff", False)
        self.blockfinder_window.lift()
        self.blockfinder_window.focus_force()

        screen_width = self.blockfinder_window.winfo_screenwidth()
        screen_height = self.blockfinder_window.winfo_screenheight()
        wwidth = self.blockfinder_window.winfo_width()
        x = (screen_width - (500 if wwidth == 1 else wwidth)) // 2
        y = (screen_height - self.blockfinder_window.winfo_height()) // 3
        self.blockfinder_window.geometry(f"+{x}+{y}")

        # Make the app responsive
        self.blockfinder_window.columnconfigure(index=0, weight=1)
        self.blockfinder_window.columnconfigure(index=1, weight=3)
        self.blockfinder_window.rowconfigure(index=0, weight=1)
        self.blockfinder_window.rowconfigure(index=1, weight=1)
        self.blockfinder_window.rowconfigure(index=2, weight=1)
        self.blockfinder_window.rowconfigure(index=3, weight=1)

        # Create control variables
        self.block_col_name = tk.StringVar(self.blockfinder_window, value="-")
        self.layer_col_name = tk.StringVar(self.blockfinder_window, value="-")

        possible_col_locations = []
        possible_col_names = []
        location = 0
        for col in self.data:
            if (
                    len(col['unique_vals']) > 0
                    and len(col['data']) > len(col['unique_vals'])
                    and len(', '.join(col['unique_vals'][0:3])) > 0
            ):
                possible_col_locations.append(location)
                possible_col_names.append(str(location+1)+': '+', '.join(col['unique_vals'][0:3]))
            location +=1


        # --------------------------------------------------------------
        # Header
        # --------------------------------------------------------------

        self.lbl_head = customtkinter.CTkLabel(self.blockfinder_window,
                                               text='Please select the name of your block(s) and layer(s) from the dropdown boxes.',
                                               justify='center')
        self.lbl_head.grid(row=0, column=0, padx=(10, 10), pady=(10, 5), sticky="nsew", columnspan=2)

        # --------------------------------------------------------------
        # Text Vaules (layer, block, id, type)
        # --------------------------------------------------------------

        self.lbl_layer= customtkinter.CTkLabel(self.blockfinder_window,
                                                text='Layer Names:',
                                                justify='left')
        self.lbl_layer.grid(row=1, column=0, padx=(10, 5), pady=(5, 5), sticky="ew")

        self.dpd_layer = customtkinter.CTkComboBox(self.blockfinder_window,
                                                   values=possible_col_names,
                                                   variable=self.layer_col_name,
                                                   state="readonly",
                                                   command=self.layer_dpd_callback)
        self.dpd_layer.grid(row=1, column=1, padx=(5, 10), pady=(5, 5), sticky="ew")

        self.lbl_block = customtkinter.CTkLabel(self.blockfinder_window,
                                                text='Block Names:',
                                                justify='left')
        self.lbl_block.grid(row=2, column=0, padx=(10, 5), pady=(5, 5), sticky="ew")

        self.dpd_block = customtkinter.CTkComboBox(self.blockfinder_window,
                                                   values=possible_col_names,
                                                   variable=self.block_col_name,
                                                   state="readonly",
                                                   command=self.block_dpd_callback)
        self.dpd_block.grid(row=2, column=1, padx=(5, 10), pady=(5, 5), sticky="ew")

        # --------------------------------------------------------------
        # Button: Open file
        # --------------------------------------------------------------

        self.btn_open = customtkinter.CTkButton(self.blockfinder_window,
                                                text="Next",
                                                command=self.block_selected_callback,
                                                corner_radius= self.ctk_btn_corner_radius,
                                                fg_color=self.ctk_momentry_btn_fg,
                                                hover_color=self.ctk_momentry_btn_hover,
                                                state='disabled')
        self.btn_open.grid(row=3, column=0, columnspan=2, padx=(10,10), pady=(5,10), sticky="ns")

        # --------------------------------------------------------------
        # Admin/Mainloop
        # --------------------------------------------------------------

        screen_width = self.blockfinder_window.winfo_screenwidth()
        screen_height = self.blockfinder_window.winfo_screenheight()
        wwidth = self.blockfinder_window.winfo_width()
        x = (screen_width - (500 if wwidth == 1 else wwidth)) // 2
        y = (screen_height - self.blockfinder_window.winfo_height()) // 3
        self.blockfinder_window.geometry(f"+{x}+{y}")

        return None

    def block_dpd_callback(self, event):
        self.block_selected = True

        self.block_column_no = int(self.dpd_block.get().split(':')[0])-1

        if self.block_selected and self.layer_selected and self.block_column_no != self.layer_column_no:
            self.btn_open.configure(state='normal')

    def layer_dpd_callback(self, event):
        self.layer_selected = True

        self.layer_column_no = int(self.dpd_layer.get().split(':')[0])-1

        if self.block_selected and self.layer_selected and self.block_column_no != self.layer_column_no:
            self.btn_open.configure(state='normal')

    def block_selected_callback(self):
        self.block_column_no = int(self.dpd_block.get().split(':')[0])-1
        self.layer_column_no = int(self.dpd_layer.get().split(':')[0])-1

        self.blockfinder_window.destroy()

        self.layer_column = self.data.copy().pop(self.layer_column_no)
        self.block_column = self.data.copy().pop(self.block_column_no)

        logger.info('    - Importer: Block/Layers Selected')

        self.points = []
        for block in self.block_column['unique_vals']:
            self.filtered_data = self.filter_to_cols(block)

            if self.block_column_no < self.layer_column_no:
                self.layer_column = self.filtered_data.pop(self.layer_column_no)
                self.block_column = self.filtered_data.pop(self.block_column_no)
            else:
                self.block_column = self.filtered_data.pop(self.block_column_no)
                self.layer_column = self.filtered_data.pop(self.layer_column_no)

            self.keyfinder_gui(block)

            self.parent.wait_window(self.keyfinder_window)

            for i in range(len(self.layer_column['data'])):
                sorted_data = {}
                sorted_data['layer'] = self.layer_column['data'][i]
                sorted_data['block'] = self.block_column['data'][i]
                sorted_data['id'] = self.pntid_column['data'][i]
                sorted_data['type'] = self.ptype_column['data'][i]
                sorted_data['x'] = float(self.pontx_column['data'][i])
                sorted_data['y'] = float(self.ponty_column['data'][i])
                self.points.append(RiggingPoint(sorted_data=sorted_data))

            logger.info('    - Importer: Processed block "%s" -  %s points', block,len(self.points))

    def filter_to_cols(self, block_name=None):
        row_location = 0
        filtered_data = []
        for i in self.data:
            filtered_data.append({'data':[],
                     'unique_vals':[],
                     'type':None,
                     'max':None,
                     'min':None})
        for block in self.block_column['data']:
            if (block == block_name if block_name else True):
                col_location = 0
                for col in self.data:
                    filtered_data[col_location]['data'].append(col['data'][row_location])
                    col_location += 1
            row_location += 1

        for col in filtered_data:
            try:
                min_val = None
                for line in col['data']:
                    x = float(line.replace(' ', ''))
                    if not min_val:
                        min_val = x
                        max_val = x
                    else:
                        min_val = min(min_val, x)
                        max_val = max(max_val, x)
            except Exception as e:
                # print(str(type(e).__name__)+' encountered.'+ traceback.format_exc() + '\n' + str(e))
                col['type'] = 'text'
                for line in col['data']:
                    if line not in col['unique_vals']:
                        col['unique_vals'].append(line)
            else:
                col['type'] = 'numeric'
                col['max'] = max_val
                col['min'] = min_val

        return filtered_data

    def keyfinder_gui(self, block_name):
        # Start the window
        # self.keyfinder_window = customtkinter.CTk()
        self.keyfinder_window = customtkinter.CTkToplevel(self.parent)
        self.keyfinder_window.title('Import CSV Step 2 - attributes for "'+block_name+'"')
        self.keyfinder_window.option_add("*tearOff", False)
        self.keyfinder_window.lift()
        self.keyfinder_window.focus_force()

        screen_width = self.keyfinder_window.winfo_screenwidth()
        screen_height = self.keyfinder_window.winfo_screenheight()
        wwidth = self.keyfinder_window.winfo_width()
        x = (screen_width - (500 if wwidth == 1 else wwidth)) // 2
        y = (screen_height - self.keyfinder_window.winfo_height()) // 3
        self.keyfinder_window.geometry(f"+{x}+{y}")

        # Make the app responsive
        self.keyfinder_window.columnconfigure(index=0, weight=1)
        self.keyfinder_window.columnconfigure(index=1, weight=3)
        self.keyfinder_window.rowconfigure(index=0, weight=1)
        self.keyfinder_window.rowconfigure(index=1, weight=1)
        self.keyfinder_window.rowconfigure(index=2, weight=1)
        self.keyfinder_window.rowconfigure(index=3, weight=1)
        self.keyfinder_window.rowconfigure(index=4, weight=1)
        self.keyfinder_window.rowconfigure(index=5, weight=1)
        self.keyfinder_window.rowconfigure(index=6, weight=1)
        self.keyfinder_window.rowconfigure(index=7, weight=1)

        # Create control variables
        self.tk_dpd_id = tk.StringVar(self.keyfinder_window, value="-")
        self.tk_dpd_type = tk.StringVar(self.keyfinder_window, value="-")
        self.tk_dpd_x = tk.StringVar(self.keyfinder_window, value="-")
        self.tk_dpd_y = tk.StringVar(self.keyfinder_window, value="-")

        possible_col_locations = []
        possible_col_names = []
        possible_col_names_txt = []
        possible_col_names_num = []
        location = 0
        for col in self.filtered_data:
            if (
                    #len(col['unique_vals']) > 0 and
                    len(', '.join(col['unique_vals'][0:3])) > 0 or
                    col['max']
                ):
                possible_col_locations.append(location)
                possible_col_names.append(str(location+1)+': '+', '.join(col['unique_vals'][0:3]))
                if col['type'] == 'numeric':
                    possible_col_names_num.append(str(location+1)+': Max='+mm_to_m(col['max'])+' Min='+mm_to_m(col['min']))
                if col['type'] == 'text':
                    possible_col_names_txt.append(str(location+1)+': '+', '.join(col['unique_vals'][0:3]))
            location +=1

        # --------------------------------------------------------------
        # Header
        # --------------------------------------------------------------

        self.lbl_head = customtkinter.CTkLabel(self.keyfinder_window,
                                               text='Please identify the other attributes using the dropdown boxes below',
                                               justify='center')
        self.lbl_head.grid(row=0, column=0, padx=(10, 10), pady=(10, 5), sticky="nsew", columnspan=2)

        # --------------------------------------------------------------
        # Text Vaules (id, type)
        # --------------------------------------------------------------

        self.lbl_head = customtkinter.CTkLabel(self.keyfinder_window,
                                               text='Text Attributes',
                                               justify='center')
        self.lbl_head.grid(row=1, column=0, padx=(10, 10), pady=(5, 5), sticky="nsew", columnspan=2)



        self.lbl_id = customtkinter.CTkLabel(self.keyfinder_window,
                                                text='Point ID:',
                                                justify='left')
        self.lbl_id.grid(row=2, column=0, padx=(10, 5), pady=(5, 5), sticky="ew")

        self.dpd_id = customtkinter.CTkComboBox(self.keyfinder_window,
                                                   values=possible_col_names_txt,
                                                   variable=self.tk_dpd_id,
                                                   state="readonly",
                                                   command=self.pntid_dpd_callback)
        self.dpd_id.grid(row=2, column=1, padx=(5, 10), pady=(5, 5), sticky="ew")



        self.lbl_type = customtkinter.CTkLabel(self.keyfinder_window,
                                                text='Point Type:',
                                                justify='left')
        self.lbl_type.grid(row=3, column=0, padx=(10, 5), pady=(5, 5), sticky="ew")

        self.dpd_type = customtkinter.CTkComboBox(self.keyfinder_window,
                                                   values=possible_col_names_txt,
                                                   variable=self.tk_dpd_type,
                                                   state="readonly",
                                                   command=self.ptype_dpd_callback)
        self.dpd_type.grid(row=3, column=1, padx=(5, 10), pady=(5, 5), sticky="ew")

        # --------------------------------------------------------------
        # Numeric Vaules (X, Y)
        # --------------------------------------------------------------

        self.lbl_head = customtkinter.CTkLabel(self.keyfinder_window,
                                               text='Numeric Attributes',
                                               justify='center')
        self.lbl_head.grid(row=4, column=0, padx=(10, 10), pady=(5, 5), sticky="nsew", columnspan=2)



        self.lbl_x = customtkinter.CTkLabel(self.keyfinder_window,
                                                text='Point X (SL-SR):',
                                                justify='left')
        self.lbl_x.grid(row=5, column=0, padx=(10, 5), pady=(5, 5), sticky="ew")

        self.dpd_x = customtkinter.CTkComboBox(self.keyfinder_window,
                                                   values=possible_col_names_num,
                                                   variable=self.tk_dpd_x,
                                                   state="readonly",
                                                   command=self.pontx_dpd_callback)
        self.dpd_x.grid(row=5, column=1, padx=(5, 10), pady=(5, 5), sticky="ew")



        self.lbl_y = customtkinter.CTkLabel(self.keyfinder_window,
                                                text='Point Y (US-DS):',
                                                justify='left')
        self.lbl_y.grid(row=6, column=0, padx=(10, 5), pady=(5, 5), sticky="ew")

        self.dpd_y = customtkinter.CTkComboBox(self.keyfinder_window,
                                                   values=possible_col_names_num,
                                                   variable=self.tk_dpd_y,
                                                   state="readonly",
                                                   command=self.ponty_dpd_callback)
        self.dpd_y.grid(row=6, column=1, padx=(5, 10), pady=(5, 5), sticky="ew")

        # --------------------------------------------------------------
        # Button: Open file
        # --------------------------------------------------------------

        self.btn_next = customtkinter.CTkButton(self.keyfinder_window,
                                                text="Next",
                                                command=self.key_selected_callback,
                                                corner_radius= self.ctk_btn_corner_radius,
                                                fg_color=self.ctk_momentry_btn_fg,
                                                hover_color=self.ctk_momentry_btn_hover,
                                                state='disabled')
        self.btn_next.grid(row=7, column=0, columnspan=2, padx=(10,10), pady=(5,10), sticky="ns")

        # --------------------------------------------------------------
        # Admin/Mainloop
        # --------------------------------------------------------------

        screen_width = self.keyfinder_window.winfo_screenwidth()
        screen_height = self.keyfinder_window.winfo_screenheight()
        wwidth = self.keyfinder_window.winfo_width()
        x = (screen_width - (500 if wwidth == 1 else wwidth)) // 2
        y = (screen_height - self.keyfinder_window.winfo_height()) // 3
        self.keyfinder_window.geometry(f"+{x}+{y}")

        # self.keyfinder_window.mainloop()

    def pntid_dpd_callback(self, event):
        self.pntid_selected = True

        self.pntid_column_no = int(self.dpd_id.get().split(':')[0])-1

        if (self.pntid_selected and
                self.ptype_selected and
                self.pontx_selected and
                self.ponty_selected and
                self.pntid_column_no != self.ptype_column_no):
            self.btn_next.configure(state='normal')

    def ptype_dpd_callback(self, event):
        self.ptype_selected = True

        self.ptype_column_no = int(self.dpd_type.get().split(':')[0])-1

        if (self.pntid_selected and
                self.ptype_selected and
                self.pontx_selected and
                self.ponty_selected and
                self.pntid_column_no != self.ptype_column_no):
            self.btn_next.configure(state='normal')

    def pontx_dpd_callback(self, event):
        self.pontx_selected = True

        self.pontx_column_no = int(self.dpd_x.get().split(':')[0])-1

        if (self.pntid_selected and
                self.ptype_selected and
                self.pontx_selected and
                self.ponty_selected and
                self.pontx_column_no != self.ponty_column_no):
            self.btn_next.configure(state='normal')

    def ponty_dpd_callback(self, event):
        self.ponty_selected = True

        self.ponty_column_no = int(self.dpd_y.get().split(':')[0])-1

        if (self.pntid_selected and
                self.ptype_selected and
                self.pontx_selected and
                self.ponty_selected and
                self.pontx_column_no != self.ponty_column_no):
            self.btn_next.configure(state='normal')

    def key_selected_callback(self):
        self.pntid_column_no = int(self.dpd_id.get().split(':')[0])-1
        self.ptype_column_no = int(self.dpd_type.get().split(':')[0])-1
        self.pontx_column_no = int(self.dpd_x.get().split(':')[0])-1
        self.ponty_column_no = int(self.dpd_y.get().split(':')[0])-1

        self.pntid_column = self.filtered_data.copy().pop(self.pntid_column_no)
        self.ptype_column = self.filtered_data.copy().pop(self.ptype_column_no)
        self.pontx_column = self.filtered_data.copy().pop(self.pontx_column_no)
        self.ponty_column = self.filtered_data.copy().pop(self.ponty_column_no)

        self.keyfinder_window.destroy()

        logger.info('    - Importer: Attributes Selected')

    def return_points(self):

        self.data = parse_csv_to_cols(self.filename)
        misc = self.block_finder_gui()
        self.parent.wait_window(self.blockfinder_window)

        logger.info('    - Importer: Returning %s points', len(self.points))
        return self.points

class ScrollableCheckBoxFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, **kwargs):
        super().__init__(master, **kwargs)
        super().pack(expand=True, fill='both')

        self.ctk_momentry_btn_fg = ('green4', 'green4')  # (light,dark)
        self.ctk_momentry_btn_hover = ('dark green', 'dark green')
        self.ctk_btn_corner_radius = 10

        self.command = command
        self.checkbox_list = []
        for i, item in enumerate(item_list):
            self.add_item(item)

    def add_item(self, item, state=False):
        checkbox_state = tk.BooleanVar(self, state)
        checkbox = customtkinter.CTkCheckBox(self, text=item, variable=checkbox_state,
                                                  fg_color=self.ctk_momentry_btn_fg,
                                                  hover_color=self.ctk_momentry_btn_hover)
        if self.command is not None:
            checkbox.configure(command=self.command)
        checkbox.grid(row=len(self.checkbox_list), column=0, pady=(0, 10), sticky="w")
        self.checkbox_list.append(checkbox)

    def update_item(self, item, state=False):
        checkbox_state = tk.BooleanVar(self, state)
        for checkbox in self.checkbox_list:
            if item == checkbox.cget("text"):
                checkbox.configure(variable=checkbox_state)
                return

    def remove_item(self, item):
        for checkbox in self.checkbox_list:
            if item == checkbox.cget("text"):
                checkbox.destroy()
                self.checkbox_list.remove(checkbox)
                return

    def clear_all(self):
        for checkbox in self.checkbox_list:
            checkbox.destroy()
        self.checkbox_list = []

    def get(self, item):
        for checkbox in self.checkbox_list:
            if item == checkbox.cget("text"):
                return checkbox.cget("variable")

    def get_checked_items(self):
        return [checkbox.cget("text") for checkbox in self.checkbox_list if checkbox.get() == 1]

class FloatEntry(customtkinter.CTkEntry):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        vcmd = (self.register(self.validate),'%P')
        self.configure(validate="all", validatecommand=vcmd, justify='center')

    def validate(self, text):
        if (
            all(char in "0123456789.m" for char in text) and  # all characters are valid
            "-" not in text[1:] and # "-" is the first character or not present
            "m" not in text[:-1] and # "-" is the first character or not present
            text.count(".") <= 1): # only 0 or 1 periods
                return True
        else:
            return False

# --------------------------------------------------------------
# Testing
# --------------------------------------------------------------

def test(file):
    logger.info('Running test function with file: [%s]', file)
    points = import_csv(file)
    obj = RiggingPlot(points)
    obj.save_to_xlsx(file.split('.tx')[0]+".xlsx",
                     stagesize = (18290, 9750),
                     delay_split=None,
                     pa_mode=False)
    subprocess.call(('open', file.split('.tx')[0]+".xlsx"))

if __name__ == "__main__":
    # test(file)
    try:
        AutoMarkoutGUI()
    except Exception as e:
        tk.messagebox.showerror(title='Error',
                                message=str(type(e).__name__) + ' encountered. AutoSV has crashed.',
                                detail=e) # if self._full_trace == False else traceback.format_exc() + '\n' + str(e))