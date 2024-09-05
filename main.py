import os
import glob
from Tools import *
from docxtpl import DocxTemplate

class DocxGenerator:

    def __init__(self, nom_mission):

        self.nom_mission = nom_mission
        self.source_file = None
        self.word_template_paths = None
        self.word_template_avec = None
        self.word_template_sans = None
        self.output_folder = None

    def set_paths(self):
        """
        Set file paths based on the specified mission.
        """
        if self.nom_mission == 'AG4':
            self.source_file = "Templates_AG4/Tableau automatisation AG de réalisation scission SCA et SCF.xlsx"
            self.word_template_paths = ["Templates_AG4/SCF_PV AG 4 trame v04.09.2023.docx"]
            self.output_folder = "Templates_AG4/Output_AG4"
        if self.nom_mission == 'AG4_Annexe':
            self.source_file = "Templates_AG4/Tableau automatisation AG de réalisation scission SCA et SCF.xlsx"
            self.output_folder = "Templates_AG4/Output_AG4_Annexes"
        elif self.nom_mission == 'AG3':
            self.source_file = "Templates_AG2_AG3/AG3/Tableau d'automatisation AG 3 envoi TTT v14.06.2023.xlsx"
            self.word_template_paths = ["Templates_AG2_AG3/AG3/PV trame AG 3 v05.06.2023.docx"]
            self.output_folder = "Templates_AG2_AG3/Output_AG3"
        elif self.nom_mission == 'AG2':
            self.source_file = "Templates_AG2_AG3/AG2/Tableau d'automatisation AG 2 envoi TTT v14.06.2023.xlsx"
            self.word_template_avec = os.path.abspath("Templates_AG2_AG3/AG2/PV AG 2 trame avec distribution de réserves v19.05.2023.docx")
            self.word_template_sans = os.path.abspath("Templates_AG2_AG3/AG2/PV AG 2 trame sans distribution de réserves v05.06.2023.docx")
            self.output_folder = "Templates_AG2_AG3/Output_AG2"
        elif self.nom_mission == 'AG1':
            self.source_file = "Templates_AG1/Tableau d'automatisation AG 1 v21.04.2023.xlsx"
            self.word_template_paths = glob.glob("Templates_AG1/AG1/*.docx")
            self.output_folder = "Templates_AG1/Output_AG1"



    def run(self):
        """
        Run the DocxGenerator for the specified mission.
        """
        self.set_paths()
        df = input(self.source_file, self.nom_mission)

        if self.nom_mission != 'AG4_Annexe':
            clear_output_folder(self.output_folder)
        output(df, self.word_template_paths, self.nom_mission, self.output_folder,self.word_template_sans,self.word_template_avec)


if __name__ == '__main__':
    docx_generator = DocxGenerator('AG2')
    docx_generator.run()
