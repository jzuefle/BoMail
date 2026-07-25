import pandas as pd
import win32com.client
import mimetypes
import os
from pathlib import Path
from create_outlook_mail import create_outlook_mail

class import_pice_table:

    platics = ['PTFE', 'PP-H', 'POM schwarz', 'POM weiß', 'PE-LD']
    aluminium = ['3.3206 Al Mg Si 0,5', '3.3535 AlMg3', '3.1645 AlCuMgPb', '3.3547 AlMg4,5Mn', 'AlMgSi', 'Al']
    copper = ['2.0401 CuZn39 Pb 3', '2.0321 CuZn37']
    stainless = ['VA', 'V2A']
    purchase = ['Kaufteil']
    federkontakte = ['2.1285 CuCoBe']
    undefiend = []
    to_be_build = 1
    
    def __init__(self, file_with_bill=r'C:\Users\J.Züfle\OneDrive - FS Antennentechnik GmbH\Desktop\S23011-44s-02_eMail.xlsx'):
        file_extension = file_with_bill.split('.')[-1].lower()
        if file_extension == "iam":
            self.df = self.import_iventor_bill(file_with_bill)
        else:
            self.df = pd.read_excel(file_with_bill, )
        self.df = self.sort_table(self.df)
        self.toOrder = self.purchase
        self.toOrderString = "Alu"
        
    def get_type_by_extension(file_path: str) -> str:
        mime_type, _ = mimetypes.guess_type(file_path)
        return mime_type or "Unknown type"
        
    def import_iventor_bill(self, inventor_file_iam):
        inventor_app = win32com.client.Dispatch("Inventor.Application")
        doc = inventor_app.Documents.Open(inventor_file_iam)
        doc.DocumentType
        bom = doc.ComponentDefinition.BOM
        bom.StructuredViewFirstLevelOnly = False
        bom.StructuredViewEnabled = True
        bom.StructuredViewEnabled
        bom.PartsOnlyViewEnabled = True
        # bom_view = bom.BOMViews.Item(1)
        bom_table = pd.DataFrame(columns=[
            "Objekt",
            "ANZAHL",
            "Bauteilnummer",
            "Beschreibung",
            "Material",
            "Bemerkung",
            "Halbzeug",
            "Komponententyp",
            "Stücklistenstruktur",
            "Dateipfad"
            ])
        
        for bom_view in bom.BOMViews:
            print(f"\n=== BOM View: {bom_view.Name} ===")
            if bom_view.Name == 'Nur Bauteile' or bom_view.Name == 'Parts Only':
                pass
            else:
                continue
            for row in bom_view.BOMRows:                
                bom_table.loc[row.ItemNumber, "Objekt"] = row.ItemNumber
                bom_table.loc[row.ItemNumber, "ANZAHL"] = row.ItemQuantity
                comp_def = row.ComponentDefinitions.Item(1)
                props = comp_def.Document.PropertySets.Item("Design Tracking Properties")
                bom_table.loc[row.ItemNumber, "Bauteilnummer"] = props.Item("Part Number").Value
                bom_table.loc[row.ItemNumber, "Beschreibung"] = props.Item("Description").Value
                bom_table.loc[row.ItemNumber, "Material"] = props.Item("Material").Value
                try: bom_table.loc[row.ItemNumber, "Bemerkung"] = comp_def.Document.PropertySets.Item("Inventor User Defined Properties").Item("Bemerkung").Value
                except: bom_table.loc[row.ItemNumber, "Bemerkung"] = ""
                try: bom_table.loc[row.ItemNumber, "Halbzeug"] = comp_def.Document.PropertySets.Item("Inventor User Defined Properties").Item("Halbzeug").Value
                except: bom_table.loc[row.ItemNumber, "Halbzeug"] = ""
                structure_map = {
                    51969: "Normal",
                    51970: "Phantom",
                    51971: "Reference",
                    51972: "Gekauft",
                    51973: "Inseparable"
                }
                bom_table.loc[row.ItemNumber, "Komponententyp"] = "Bauteil"
                bom_table.loc[row.ItemNumber, "Stücklistenstruktur"] = structure_map.get(row.BOMStructure, str(row.BOMStructure))
                bom_table.loc[row.ItemNumber, "Dateipfad"] = row.ComponentDefinitions.Item(1).Document.FullFileName
                
        return bom_table
    
    def sort_table(self, table_to_sort):
        sorted_table = pd.DataFrame(columns=[
            "Baugruppenposition",
            "Bauteilnummer",
            "Anzahl",
            "Bezeichnung",
            "Bemerkung",
            "Material",
            "Halbzeug",
            "CAD_File",
            "PDF_File",
            "Kategorie",
            "Lagerwert"])
        try: sorted_table.Baugruppenposition = table_to_sort['Objekt']
        except: None
        sorted_table.Bauteilnummer = table_to_sort['Bauteilnummer']
        try: sorted_table.Anzahl = table_to_sort['ANZAHL']
        except: None
        try: sorted_table.Bezeichnung = table_to_sort[sorted(set(table_to_sort) & set(['Bezeichnung', 'Beschreibung']))[0]]
        except: None
        if any(table_to_sort.columns == 'Bemerkung'):
            sorted_table.Bemerkung = table_to_sort['Bemerkung']
        else:
            sorted_table.Bemerkung
        if any(table_to_sort['Material']):
            sorted_table.Material = table_to_sort['Material']
        if any(table_to_sort.columns == 'Halbzeug'):
            sorted_table.Halbzeug = table_to_sort['Halbzeug']
        else:
            sorted_table.Halbzeug
        sorted_table.CAD_File
        sorted_table.PDF_File
        if any(table_to_sort.columns == 'Komponententyp'):
            sorted_table.Komponententyp = table_to_sort['Komponententyp']
        if any(table_to_sort.columns == 'Kategorie'):
            sorted_table.Kategorie = table_to_sort['Kategorie']
        else:
            sorted_table.Kategorie
            for row, item in sorted_table.iterrows():
                if table_to_sort['Stücklistenstruktur'][row] == 'Gekauft':
                    sorted_table.loc[row, 'Kategorie'] = "Kaufteil"
                elif item.Bauteilnummer[:4] == "Igus" or item.Bauteilnummer[:5] == "Ensat" or item.Bauteilnummer[:7] == "Spinner" or item.Bauteilnummer[:4] == "Wago" or item.Bauteilnummer[:4] == "Murr":
                    sorted_table.loc[row, 'Kategorie'] = "Kaufteil"
                elif item.Bauteilnummer[8] == 'S':
                    sorted_table.loc[row, 'Kategorie'] = "Fertigungsteil"
                elif item.Bauteilnummer[8] == 'E':
                    sorted_table.loc[row, 'Kategorie'] = "Fertigungsteil"
                else:
                    sorted_table.loc[row, 'Kategorie'] = "Kaufteil"
                    # G2="Baugruppe";G2;WENN(TEIL(C2;9;1)="E";"Fertigungsteil";WENN(ODER(H2="Gekauft";
        if any(table_to_sort.columns == 'Lagerwert'):
            sorted_table.Lagerwert = table_to_sort['Lagerwert']
        else:
            sorted_table.Lagerwert
        
        #if table_to_sort:
        #    sorted_table = table_to_sort
                
        return sorted_table
        
    def change_bouild_number(self, to_be_build=1):
        self.to_be_build = to_be_build
        self.correct_table()
    
    def correct_table(self):
        self.df['Zu beschaffen'] = self.df.Anzahl * self.to_be_build - pd.to_numeric(self.df.Lagerwert, errors='coerce').fillna(0)
        # self.df['Zu beschaffen'] = self.df['Benötigt pro Antenne'] * self.to_be_build - pd.to_numeric(self.df.Lagerwert, errors='coerce').fillna(0)
        self.df['STCK'] = self.df['Zu beschaffen'].astype(int)
    
    def find_all_files(self, directory, file_extension):
        all_files = []
        for file_path in Path(directory).rglob(f"*{file_extension}"):
            if file_path.is_file():
                all_files.append(str(file_path))
        return all_files
    
    def find_all_drawings(self, directory=r"W:\KONSTRUKTION-3D"):
        allPDFs = []
        allPDFs.extend(self.find_all_files(directory, file_extension='pdf'))
        # allPDFs.extend(self.find_all_files(directory=r"C:\Users\J.Züfle\OneDrive - FS Antennentechnik GmbH\Projekte\4419 Beamtrail S12018-1 AMS\S12018-01b_PDF", file_extension='pdf'))
        # allPDFs.extend(self.find_all_files(directory=r"W:\KONSTRUKTION-3D\PRODUKTE\S23011-44x-Mast-mit-Wagen", file_extension='pdf'))
        # allPDFs.extend(self.find_all_files(directory=r"W:\KONSTRUKTION-3D\PRODUKTE\S23011-4-1x-Vierergruppe", file_extension='pdf'))
        # allPDFs.extend(self.find_all_files(directory=r"C:\Users\J.Züfle\OneDrive - FS Antennentechnik GmbH\Projekte\4419 Beamtrail S12018-1 AMS", file_extension='pdf'))
                                           
        allSTPs = []
        allSTPs.extend(self.find_all_files(directory, file_extension='stp'))
        # allSTPs.extend(self.find_all_files(directory=r"W:\KONSTRUKTION-3D\PRODUKTE\S23011-44x-Mast-mit-Wagen", file_extension='stp'))
        # allSTPs.extend(self.find_all_files(directory=r"W:\KONSTRUKTION-3D\PRODUKTE\S23011-4-1x-Vierergruppe", file_extension='stp'))
        
        fileLocation = pd.DataFrame({'Bauteilnummer': [''], 'PDF-Pfad': [''], 'STP-Pfad': ['']})

        for pdfToAdd in allPDFs:
            partnumber = pdfToAdd.split('\\')[-1]
            partnumber = '.'.join(partnumber.split('.')[:-1])
            if not(any(fileLocation['Bauteilnummer'] == partnumber)):
                fileLocation.loc[len(fileLocation), 'Bauteilnummer'] = partnumber
            if any(fileLocation[fileLocation['Bauteilnummer'] == partnumber]['PDF-Pfad'].isna()):
                fileLocation.loc[fileLocation['Bauteilnummer'] == partnumber, 'PDF-Pfad'] = pdfToAdd

        for stpToAdd in allSTPs:
            partnumber = stpToAdd.split('\\')[-1]
            partnumber = '.'.join(partnumber.split('.')[:-1])
            if not(any(fileLocation['Bauteilnummer'] == partnumber)):
                fileLocation.loc[len(fileLocation), 'Bauteilnummer'] = partnumber
            if any(fileLocation[fileLocation['Bauteilnummer'] == partnumber]['STP-Pfad'].isna()):
                fileLocation.loc[fileLocation['Bauteilnummer'] == partnumber, 'STP-Pfad'] = stpToAdd
        
        return fileLocation
    
    def create_list_to_order(self, material_to_order):
        if material_to_order == 'Aluminium':
            toOrder = self.aluminium
        elif material_to_order == 'Kunststoff':
            toOrder = self.platics
        elif material_to_order == 'Kupfer':
            toOrder = self.copper
        elif material_to_order == 'Edelstahl':
            toOrder = self.stainless
        elif material_to_order == 'Federkontakte':
            toOrder = self.federkontakte
        elif material_to_order == 'Kaufteile':
            toOrder = self.purchase
        else:
            toOrder = self.undefiend
        toOrderString = material_to_order

        if toOrder == self.purchase:
            toFillInTable = self.df[self.df.Kategorie == self.purchase[0]][self.df['Zu beschaffen']>0][['STCK', 'Bauteilnummer', 'Bezeichnung', 'Halbzeug', 'Bemerkung']]
        elif toOrder == self.undefiend:
            all_options = self.aluminium + self.platics + self.copper + self.stainless + self.purchase
            toFillInTable = self.df[self.df.Kategorie == 'Fertigungsteil'][~self.df.Material.isin(all_options)][self.df['Zu beschaffen']>0][['STCK', 'Bauteilnummer', 'Bezeichnung', 'Halbzeug', 'Bemerkung']]
        elif toOrder == self.federkontakte:            
            toFillInTable = self.df[self.df.Material.isin(toOrder)][self.df['Zu beschaffen']>0][['STCK', 'Bauteilnummer', 'Bezeichnung', 'Halbzeug', 'Bemerkung']]
        else:
            toFillInTable = self.df[self.df.Kategorie=='Fertigungsteil'][self.df.Material.isin(toOrder)][self.df['Zu beschaffen']>0][['STCK', 'Bauteilnummer', 'Bezeichnung', 'Halbzeug', 'Bemerkung']]
        toFillInTable = toFillInTable.fillna('')
        
        return toFillInTable
    
    def all_drawings_to_order(self, material_to_order):
        fileLocation = self.find_all_drawings()
        toFillInTable = self.create_list_to_order(material_to_order)

        addPdf = []
        for addedFile in toFillInTable['Bauteilnummer']:
            if any(fileLocation['Bauteilnummer'] == addedFile) and not(any(fileLocation[fileLocation['Bauteilnummer'] == addedFile]['PDF-Pfad'].isna())):
                addPdf.append(fileLocation.loc[fileLocation['Bauteilnummer'] == addedFile, 'PDF-Pfad'].iloc[0])
        
        addStp = []
        for addedFile in toFillInTable['Bauteilnummer']:
            if any(fileLocation['Bauteilnummer'] == addedFile) and not(any(fileLocation[fileLocation['Bauteilnummer'] == addedFile]['STP-Pfad'].isna())):
                addStp.append(fileLocation.loc[fileLocation['Bauteilnummer'] == addedFile, 'STP-Pfad'].iloc[0])

        addFile = addPdf + addStp
        return addFile
    