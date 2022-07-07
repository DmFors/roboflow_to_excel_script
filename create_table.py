import json
import csv
import os
from collections import Counter
import pandas as pd

COLUMNS = ["Number", "Name_file"]


def load_descriptions_json(filepath):
    with open(filepath, "r") as file:
        text = file.read()
    descriptions = json.loads(text)
    return descriptions


def generate_table(descriptions, classes):
    table = []

    number = 1
    for image in descriptions:
        name_file = image["image"].split(".rf.")[0].replace("_", ".")
        annotations = image["annotations"]
        amount_entities = generate_amount_entities(classes, annotations)
        row = {
            COLUMNS[0]: number,
            COLUMNS[1]: name_file,
        }
        row.update(amount_entities)
        table.append(row)
        number += 1
    return table


def generate_amount_entities(classes, annotations):
    entities = []
    for annotation in annotations:
        entity = annotation["label"]
        entities.append(entity)
    counter = dict(Counter(entities))
    amount_entities = {}
    for label in classes:
        if label in counter:
            amount_entities[label] = counter[label]
        else:
            amount_entities[label] = 0
    return amount_entities


def create_csv(filename, table, classes):
    with open(filename + ".csv", 'w') as file:
        csv_writer = csv.DictWriter(file, delimiter=',', lineterminator='\r', fieldnames=COLUMNS + classes,
                                    dialect='excel')
        csv_writer.writeheader()
        csv_writer.writerows(table)


def convert_from_csv_to_excel(filename_from, filename_to):
    csv_file = pd.read_csv(filename_from + ".csv")
    csv_file.to_excel(filename_to + ".xlsx", index=None, header=True)


def delete_file(filename):
    os.remove(filename)


def create_excel(directory_of_dataset, classes):
    images_sets = ("train",)
    annotations = []
    print("Getting annotations...")
    for images_set in images_sets:
        annotations += load_descriptions_json(directory_of_dataset + images_set + "/_annotations.createml.json")
    print("Generating table...")
    table = generate_table(annotations, classes)
    create_csv("output", table, classes)
    print("Converting to Excel...")
    convert_from_csv_to_excel("output", "report")
    print("Deleting temporal files...")
    delete_file("output.csv")
    # delete_file(directory)
    print("Excel table created!")
