from escpos.printer import Usb

"""
Prints a markout to a standard ESC/POS receipt printer.

this is my printer:
https://www.qsprinter.com/products/15.html

this is the docs
https://python-escpos.readthedocs.io/en/latest/index.html
"""

class escpos_printer:

    # Designed to be used in a context manager so that it shuts (cuts) automatically at the end

    def __init__(self, idVendor=0x0485, idProduct=0x7541, in_ep=130, out_ep=3):

        # Just sets internal variables ready for the __enter__() function

        self._idVendor = idVendor
        self._idProduct = idProduct
        self._in_ep = in_ep
        self._out_ep = out_ep

    def __enter__(self):

        # Opens the printer connection and returns this class instance

        self.printer = Usb(idVendor=self._idVendor,
                           dProduct=self._idProduct,
                           in_ep=self._in_ep,
                           out_ep=self._out_ep)

        return self

    def print_source(self, data, source_num):

        source = data['sources'][source_num]

        # Writes an AutoSV source to an ESC/POS printer
        # Written for a specific very cheap 38mm printer

        self.printer.line_spacing(spacing=0, divisor=180)

        # --------------------------------------------------------------
        # City Name
        # --------------------------------------------------------------

        self.printer.set(align='center',
                         font=0,
                         flip = False,
                         double_height= False,
                         double_width= False,
                         custom_size=True,
                         width=3,
                         height=3,
                         invert = False,
                         smooth =False,
                         bold = False)

        self.printer.text(data['project']['city']+"\n\n")

        # --------------------------------------------------------------
        # Source Name and Data
        # --------------------------------------------------------------

        self.printer.set(align='center',
                         font=0,
                         custom_size=True,
                         width=2,
                         height=2)

        self.printer.text(source['name']+"\n\n")

        self.printer.set(align='center',
                         font=1,
                         flip = False,
                         double_height= False,
                         double_width= False,
                         custom_size=True,
                         width=2,
                         height=1,
                         invert = False,
                         smooth =False,
                         bold = False)

        self.printer.text("Weight:" + str('%.2f' % source['totalWeight']).rjust(8) + " kg \n")

        self.printer.text("Top Z: " + str('%.2f' % source['position']['z']).rjust(8) + " m  \n")

        self.printer.text("Btm Z: " + str('%.2f' % source['bottom_z']).rjust(8) + " m  \n")

        self.printer.text("Site:  " + str('%.2f' % source['orientation']['x']).rjust(8) + " deg\n\n")

        # --------------------------------------------------------------
        # Bumper and pickups
        # --------------------------------------------------------------

        blocking = 4

        try:
            if 'flown' in source['config'] and 'motors' in source:
                # Bumper Cells
                # If the Bump is inverted, make the cells red
                if 'Inv' in source['bumper'] or 'inv' in source['bumper']:

                    self.printer.set(invert=True,
                         font=0,
                         custom_size=True,
                         width=2,
                         height=2)

                    self.printer.text("--INVERTED BAR--\n\n")

                    self.printer.set(invert=False,
                         font=1,
                         custom_size=True,
                         width=2,
                         height=1)

                if "+" in source['bumper']:
                    bits = source['bumper'].split("+")
                    self.printer.text(bits[0].replace('_', ' ') + " +\n" + bits[1].replace('_', ' ') + "\n")
                else:
                    self.printer.text(source['bumper'].replace('_', ' ') + "\n")

                # if "M-BUMP" in source['bumper']:
                #     blocking = 3
                # else:
                #     blocking = 4

                # Pickup Points
                if len(source['motors']) == 1:
                    self.printer.text('1 Pickup: Hole' + source['motors'][0]['pickupHole'] + "\n\n")
                else:
                    if source['motors'][1]['name'] == 'Pullback':
                        self.printer.text('Hole: ' + source['motors'][0]['pickupHole'] + " + Pullback\n\n")
                    else:
                        self.printer.text('Hole: ' + source['motors'][0]['pickupHole'] + ' / Hole: ' + source['motors'][1]['pickupHole']+ "\n\n")
            else:
                if 'bumper' in source:
                    if "+" in source['bumper']:
                        bits = source['bumper'].split("+")
                        self.printer.text(bits[0].replace('_', ' ') + " +\n" + bits[1].replace('_', ' ') + "\n")
                    else:
                        self.printer.text(source['bumper'].replace('_', ' ') + "\n")
                    if "M-BUM" in source['bumper']:
                        blocking = 3
                    else:
                        blocking = 4
                else:
                    self.printer.text("No Bumper Info\n")
                if 'flown' in source['config']:
                    self.printer.text("Fixed\n\n")
                else:
                    self.printer.text("Stacked\n\n")
        except KeyError:
            self.printer.text("No Bumper Info\n\n")

        # --------------------------------------------------------------
        # Boxes and angles
        # --------------------------------------------------------------

        self.printer.set(align='left',
                         font=1,
                         double_height=False,
                         double_width=False,
                         custom_size=True,
                         width=2,
                         height=1,
                         )
        counter=1
        for enclosure in source['enclosures']:
            print_names = enclosure.print_names(source['position']['x'], source['orientation']['z'])
            for name in print_names:
                self.printer.text(name.ljust(13) + " -" + str(enclosure.print_angle()).rjust(6) + "\n")
            if len(print_names) > 1:
                self.printer.text("\n")
            else:
                if counter == blocking:
                    self.printer.text("\n")
                    counter=1
                else:
                    counter+=1


    def __exit__(self, exc_type, exc_value, traceback):

        # Closes the printer connection

        self.printer.cut()

        self.printer.close()

def device_finder():
    import sys
    import usb.core
    # find USB devices
    dev = usb.core.find(find_all=True)
    # loop through devices, printing vendor and product ids in decimal and hex
    devices = []
    for cfg in dev:
        # sys.stdout.write(str(cfg.product) + ' Hexadecimal VendorID=' + hex(cfg.idVendor) + ' & ProductID=' + hex(cfg.idProduct) + '\n\n')
        if not ("Hub" in str(cfg.product) or "LAN" in str(cfg.product)):
            devices.append( {"product": str(cfg.product),
                      "idVendor": int(cfg.idVendor),
                      "idProduct": int(cfg.idProduct)})

    return devices

if __name__ == '__main__':

    pass