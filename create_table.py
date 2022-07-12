import json
import os
from collections import Counter
import pandas as pd

COLUMNS_FOR_EXCEL = ["Number", "Name_file", "car", "tram", "bus", "pedestrian",
                     "scooter", "bicycle", "shuttle_taxi", "trolleybus", "truck", "motorcycle"]
COLUMNS_FOR_WORD = ["Number", "Name_file", "Description"]

TRANSLATION = {"car": "машина", "tram": "трамвай", "bus": "автобус", "pedestrian": "пешеход", "scooter": "самокат",
               "bicycle": "велосипед", "shuttle_taxi": "маршрутка", "trolleybus": "троллейбус", "truck": "грузовик",
               "motorcycle": "мотоцикл"}


def load_descriptions_json(filepath):
    with open(filepath, "r") as file:
        text = file.read()
    descriptions = json.loads(text)
    return descriptions


def get_annotations(directory_of_dataset, images_sets):
    annotations = []
    for images_set in images_sets:
        annotations += load_descriptions_json(directory_of_dataset + images_set + "/_annotations.createml.json")
    return annotations


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
    return amount_entities


def generate_table_for_excel(annotations, amount_entities):
    table = []
    number = 1
    for annotation in annotations:
        image_name = annotation["image"]
        file_name = get_image_name(image_name)
        row = {
            COLUMNS_FOR_EXCEL[0]: number,
            COLUMNS_FOR_EXCEL[1]: file_name,
        }
        row.update(amount_entities[image_name])
        table.append(row)
        number += 1
    return table


def generate_table_for_word(annotations, amount_entities):
    table = []
    number = 1
    for annotation in annotations:
        image_name = annotation["image"]
        file_name = get_image_name(image_name)
        description = ""
        for entity, amount in amount_entities[image_name].items():
            description += f"{TRANSLATION[entity]} - {amount}\n"

        description = description[:-1]
        row = {
            COLUMNS_FOR_WORD[0]: number,
            COLUMNS_FOR_WORD[1]: file_name,
            COLUMNS_FOR_WORD[2]: description,
        }
        table.append(row)
        number += 1
    return table


def get_image_name(image_name):
    return image_name.split(".rf.")[0].replace("_", ".", ).replace("--------", " - копия")


def create_excel(filename, table, column_names):
    df = pd.DataFrame(table, columns=column_names)
    while True:
        try:
            writer = pd.ExcelWriter(filename + ".xlsx", engine='xlsxwriter')
        except PermissionError:
            input("Please, close Excel app and press enter...")
        else:
            break
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


def delete_file(filepath):
    os.remove(filepath)


def create_file(directory_of_dataset):
    images_sets = ("train",)
    print("Getting annotations...")
    annotations = get_annotations(directory_of_dataset, images_sets)
    print("Getting entities...")
    amount_entities = generate_amount_entities(annotations)
    print("Generating excel table for excel report...")
    table_excel_report = generate_table_for_excel(annotations, amount_entities)
    print("Generating excel table for word report...")
    table_word_report = generate_table_for_word(annotations, amount_entities)
    print("Creating excel files...")
    create_excel("for_excel_report", table_excel_report, COLUMNS_FOR_EXCEL)
    create_excel("for_word_report", table_word_report, COLUMNS_FOR_WORD)
    print("Excel tables created!")
