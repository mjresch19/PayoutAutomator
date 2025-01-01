import csv

def read_csv(file_path):
    '''
    Translate csv contents into a python datatype

    @param file_path: The path to the csv file
    @return A list of lists containing the contents of the csv file
    '''

    file_contents = []

    with open(file_path, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)

        #Skip the header
        header = next(reader)

        for row in reader:
            
            file_contents.append(row)


    return file_contents