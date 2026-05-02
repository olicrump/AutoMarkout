import sys
import ezdxf

filename1 = "/Users/olicrump/My Drive/Audio Work/20260401 Rick Astley/CAD/Old/0417 Manchester COOP (RA 26).dxf"
filename2 = "/Users/olicrump/My Drive/Audio Work/20260401 Rick Astley/CAD/Old/RA BG V1.4.dxf"

class DXF_Importer:
    def __init__(self, filename):


        self.doc = ezdxf.readfile(filename)
        self.msp = self.doc.modelspace()

    def findnames(self):
        names = {}
        for e in self.msp:
            if e.dxftype() == "INSERT":
                name = e.dxf.name
                if name not in names:
                    names[name] = 1
                else:
                    names[name] += 1
        print(names)
        #
        # # entity query for all LINE entities in modelspace
        # for e in msp.query("LINE"):
        #     print_entity(e)

        self.blockrefs = self.msp.query('INSERT[name=="My Multi Block 2025 PRACTICE"]')

        for entity in self.blockrefs:
            print(entity.dxf.name)
            for attrib in entity.attribs:
                print('"'+attrib.dxf.tag.ljust(20) + '"   "' + attrib.dxf.text+'"')

    def find_anon(self):
        # Collect all anonymous block references starting with '*U'
        anonymous_block_refs = self.msp.query('INSERT[name ? "^\*U.+"]')



        # Collect the references of the 'FLAG' block
        flag_refs = []
        for block_ref in anonymous_block_refs:
            # Get the block layout of the anonymous block
            print(block_ref.dxf.name)
            block = self.doc.blocks.get(block_ref.dxf.name)
            # Find all block references to 'FLAG' in the anonymous block
            print(block)
            misc = block.query('INSERT')
            flag_refs.extend(block.query('INSERT'))

            for entity in misc:
                print("Entity Name: " + entity.dxf.name)
                for attrib in entity.attribs:
                    print('    Attribute: "' + attrib.dxf.tag.ljust(20) + '"   "' + attrib.dxf.text + '"')

        # Evaluation example: collect all flag names.
        flag_numbers = [
            flag.get_attrib_text("LOAD")
            for flag in flag_refs
            if flag.has_attrib("LOAD")
        ]

        print(flag_refs)
        print(e.dxf.name for e in flag_refs)

if __name__ == '__main__':
    importer = DXF_Importer(filename2)
    importer.find_anon()
    # importer.findnames()