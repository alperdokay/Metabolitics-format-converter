# Author = Alper Dokay
import json
import os
import xlsxwriter

class MetaboliticsFormatConverter:
    def __init__(self, path):
        self.Path = path
    
    def convert_multiple(self):
        
        file_paths = []
        file_names = []
        for file in os.listdir(self.Path):
            if file.endswith(".json"):
                file_paths.append(os.path.join(self.Path, file))
                file_names.append(file.split('.')[0])

        workbook = xlsxwriter.Workbook('multiple_output.xlsx')

        data_worksheet = workbook.add_worksheet('data')
        data_metasheet = workbook.add_worksheet('meta')

        data_metasheet.write(0, 0, 'Study Name')
        data_metasheet.write(0, 1, '[Put study name here...]')
        data_metasheet.write(1, 0, 'Control/Healthy/Wildtype Group')
        data_metasheet.write(1, 1, '[Put group name here...]')
        data_metasheet.write(2, 0, 'Subject Id')
        data_metasheet.write(2, 1, 'Group')

        isLabelLoad = False

        data_col = 1
        meta_row, meta_col = 3, 0

        for path, name in zip(file_paths, file_names):
            data_row = 1
            with open(path, "r") as read_file:
                
                data = json.load(read_file)
                if not isLabelLoad:
                    for key in data.keys():
                        data_worksheet.write(data_row, data_col - 1, key)
                        data_row += 1
                    data_row = 1
                    isLabelLoad = True
                
                data_worksheet.write(0, data_col, name)

                for value in data.values():
                    data_worksheet.write(data_row, data_col, value)
                    data_row += 1
                
                data_metasheet.write(meta_row, meta_col, name)
                data_metasheet.write(meta_row, meta_col + 1, "[Put group name here...]")
                meta_row += 1
                data_col += 1

        workbook.close()

                

    def convert_single_file(self):
        workbook = xlsxwriter.Workbook('single_output.xlsx')

        data_worksheet = workbook.add_worksheet('data')

        data_worksheet.write(0, 1, 'Individual')
        
        with open(self.Path, "r") as read_file:

            data = json.load(read_file)
            row, col = 1, 0
            
            for key, value in data.items():
                data_worksheet.write(row, col, key)
                data_worksheet.write(row, col + 1, value)
                row += 1

        data_metasheet = workbook.add_worksheet('meta')

        data_metasheet.write(0, 0, 'Study Name')
        data_metasheet.write(0, 1, '[Put study name here...]')
        data_metasheet.write(1, 0, 'Control/Healthy/Wildtype Group')
        data_metasheet.write(1, 1, '[Put group name here...]')
        data_metasheet.write(2, 0, 'Subject Id')
        data_metasheet.write(2, 1, 'Group')
        data_metasheet.write(3, 0, 'Individual 1')
        if 'healthy' in self.Path.lower():
            data_metasheet.write(3, 1, 'healthy')
        else:
            data_metasheet.write(3, 1, 'breast cancer')

        workbook.close()

    def run(self):
        if os.path.isfile(self.Path):
            self.convert_single_file()
        else:
            self.convert_multiple()


# If you want to create a single output, give the json file directly as follows.
# converter = MetaboliticsFormatConverter('D:\Metabolitics\Excel converter\data\lung_plasma_Adenocarcinoma_6.json')
# If you want to create a multiple output, give the directory of files, it will fetch all the json files in it and convert
# converter = MetaboliticsFormatConverter('D:\Metabolitics\Excel converter\data\Study 1')
converter = MetaboliticsFormatConverter('[Give your path here!!!]')
converter.run()