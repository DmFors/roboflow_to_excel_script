import json
import csv
import os
from collections import Counter
import pandas as pd

COLUMNS = ["Number", "Name_file", "car", "tramm", "bus", "pedestrian",
           "scooter", "bicycle", "shuttle_taxi", "trolleybus", "truck", "motorcycle"]


def load_descriptions_json(filepath):
    with open(filepath, "r") as file:
        text = file.read()
    descriptions = json.loads(text)
    return descriptions


def generate_table(annotations, amount_entities):
    table = []
    number = 1
    for annotation in annotations:
        image_name = annotation["image"]
        file_name = image_name.split(".rf.")[0].replace("_", ".")
        row = {
            COLUMNS[0]: number,
            COLUMNS[1]: file_name,
        }
        row.update(amount_entities[image_name])
        table.append(row)
        number += 1
    return table


def generate_amount_entities(descriptions):
    amount_entities = {}
    for description in descriptions:
        image = description["image"]
        annotations = description["annotations"]
        entities = []
        for annotation in annotations:
            entity = annotation["label"]
            entities.append(entity)
        amount_entities[image] = dict(Counter(entities))
    # for label in labels:
    #     if label in counter:
    #         amount_entities[label] = counter[label]
    #     else:
    #         amount_entities[label] = 0
    return amount_entities


def create_csv(filename, table):
    with open(filename + ".csv", 'w') as file:
        csv_writer = csv.DictWriter(file, delimiter=',',
                                    lineterminator='\r',
                                    fieldnames=COLUMNS,
                                    dialect='excel')
        csv_writer.writeheader()
        csv_writer.writerows(table)


def convert_from_csv_to_excel(filename_from, filename_to):
    csv_file = pd.read_csv(filename_from + ".csv")
    if os.path.exists(filename_to + ".xlsx"):
        response = int(input("File report.xlsx exists, overwrite file? (1 - yes, 0 - no): "))
        if response == 1:
            while True:
                try:
                    delete_file(filename_to + ".xlsx")
                except PermissionError:
                    input("Please, close Excel app and press enter...")
                else:
                    break
        else:
            exit(0)
    csv_file.to_excel(filename_to + ".xlsx", index=None, header=True)


def delete_file(filepath):
    os.remove(filepath)


def create_excel(directory_of_dataset):
    images_sets = ("train",)
    annotations = []
    print("Getting annotations...")
    for images_set in images_sets:
        annotations += load_descriptions_json(directory_of_dataset + images_set + "/_annotations.createml.json")
    print("Generating table...")
    amount_entities = generate_amount_entities(annotations)
    table = generate_table(annotations, amount_entities)
    create_csv("output", table)
    print("Converting to Excel...")
    convert_from_csv_to_excel("output", "report")
    print("Deleting temporal files...")
    delete_file("output.csv")
    # delete_file(directory)
    print("Excel table created!")
