from roboflow import Roboflow
from create_table import create_excel

print("Welcome!")
API_KEY = input("Enter your API KEY: ")
PROJECT_ID = input("Enter your PROJECT ID: ")
VERSION = input("Enter your VERSION OF DATASET: ")

rf = Roboflow(api_key=API_KEY)
workspace = rf.workspace()
project = workspace.project(PROJECT_ID)
dataset = project.version(VERSION).download("createml")
directory = project.name.replace(' ', '-') + f"-{dataset.name}/"
classes = project.classes
create_excel(directory, classes)
