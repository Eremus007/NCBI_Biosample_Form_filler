# Imports
import os
import csv
import shutil
import pandas as pd
from pandas import DataFrame


def get_filenames(directory):
    """Iterates through a specified directory and returns a list with
    the file names.

        Parameters
        ----------
        directory : str
            The file containing the files to iterate through

        Returns
        -------
        filenames : list
            a list of strings which are the names of the files in the
            specified directory.
        """
    # iterate over files in
    # that directory
    filenames = []
    for filename in os.listdir(directory):
        file = os.path.join(directory, filename)
        # checking if it is a file
        if os.path.isfile(file):
    # split off extentions
            file_list = file.split('.')
    # remove directory path/extentions and keep file name
            filenames.append(file_list[0][14:])
    return filenames


def open_csv_doc(directory, dict_list, samplenames_list):
    """Translates an xlsx file to CSV and reads the CSV file
    then Updates the dictionaries in the argument with the data
    read from the file and returns the modified argument.

        Parameters
        ----------
        directory : str
            The file location of the xlsx file
        dict_list : list
            a list containing dictionaries with the header elements as keys
        samplenames_list : list
            a list containing the names of the sample files

        Returns
        -------
        dict_list: list
            a list of dictionaries with updates elements from the CSV file.
        """
    # Translate xlsx file into CSV file
    filecodes, code_dict = [], {}
    filenames = samplenames_list.copy()
    for num, n in enumerate(filenames):
        filecodes.append(n[0:3])
    filecodes.sort()
    read_file = pd.read_excel(
        r'Molecular_samples_all.xlsx',
        sheet_name='all_samples')
    read_file.to_csv(r'%s' % directory, header=True, index=False)
    # Write organism name to dictionary from name extracted from CSV file
    with open('%s' % directory, 'r') as file:
        csvreader = csv.reader(file, delimiter=',')
        #Skip the header
        next(csvreader)
        # update dictionary using a loop
        for num, row in enumerate(csvreader):
            if not num >= len(filenames) and filenames[num][0:3] not in code_dict.keys():
                code_dict.update({filenames[num][0:3]: row[2]})
            if not num >= len(dict_list) and not num >= len(filecodes):
                dict_list[num].update({'*sample_name': filenames[num]})
                dict_list[num].update({'*organism': code_dict.get(filenames[num][0:3])})
        return dict_list


def create_dictionary(header, samplenames_list):
    """Creates the list of dictionaries and updates some keys with
    readily acquired data

        Parameters
        ----------
        header : list
            a list of strings with the keys for the dictionaries.
        samplenames_list : list
            A list of filenames which will be updated to the
            newly created dictionary

        Returns
        -------
        data_information
            a list of dictionaries.
        """
    data_information = []
    for i in range(len(samplenames_list)):
        temp_dict = {}
        for num, element in enumerate(header):
            if num == 10:
                temp_dict.update({element: 'leaf'})
            else:
                temp_dict.update({element: ''})
        data_information.append(temp_dict)
    return data_information


def main():
    # assign directory
    directory = 'raw_sequences'
    outputfile_name = 'Filled_SRA_Plant.1.0_NeoMyristicaceae'
    filenames = get_filenames(directory)
    filenames.sort()
    header = ['*sample_name', 'sample_title', 'bioproject_accession',
              '*organism', 'isolate', 'cultivar', 'ecotype', 'age',
              'dev_stage', '*geo_loc_name', '*tissue',
              'biomaterial_provider', 'cell_line', 'cell_type',
              'collected_by', 'collection_date',
              'culture_collection', 'disease',
              'disease_stage', 'genotype',
              'growth_protocol', 'height_or_length',
              'isolation_source', 'lat_lon', 'phenotype',
              'population', 'sample_type', 'sex', 'specimen_voucher',
              'temp', 'treatment', 'description']
    directory2 = "Molecular_samples_all.csv"
    dict_list = create_dictionary(header, filenames)
    dict_list = open_csv_doc(directory2, dict_list, filenames)
    data_frame: DataFrame = pd.DataFrame(dict_list)
    # Delete target file when already present, then copy template.
    if os.path.isfile('%s.xlsx'%outputfile_name):
        os.remove('%s.xlsx'%outputfile_name)
    shutil.copy('SRA_Plant.1.0_NeotMyristicaceae.xlsx', '%s.xlsx'%outputfile_name)
    # Append DataFrame to new tsv file
    data_frame.to_csv("%s.tsv"%outputfile_name, sep='\t', index=False)
    #Append Dataframe to existing excel file
    with pd.ExcelWriter('%s.xlsx'%outputfile_name, mode='a', if_sheet_exists='replace') as writer:
        data_frame.to_excel(writer, sheet_name='Plant.1.0', header=True, index=False)
    if os.path.isfile('Molecular_samples_all.csv'):
        os.remove('Molecular_samples_all.csv')


main()
