# Main script to automate the process of generating Visio flowcharts from notes on a spreadsheet
# TODO: Add Aspose License details

import jpype
import pandas as pd
from flowchart_node import FlowchartNode

# Start JVM
if not jpype.isJVMStarted():
    jpype.startJVM()

import asposediagram
from asposediagram.api import *

OUTPUT_DIRECTORY = "../output"
INPUT_DIRECTORY = "../input"
TEMPLATE_DIRECTORY = "../templates"

# Read meeting notes spreadsheet file from input directory
input_file = f"{INPUT_DIRECTORY}/Meeting_notes.xlsx"
df = pd.read_excel(input_file)

# Generate Visio file
template_file = f"{TEMPLATE_DIRECTORY}/Basic Shapes.vss"
diagram = Diagram(template_file)

# Parse notes content to extract flowchart elements
flowchart_nodes = []

print(diagram.getMasters().isExist("Rectangle"))
print(diagram.getMasters().isExist("Process"))
print(diagram.getMasters().isExist("Diamond"))
print(diagram.getMasters().isExist("Circle"))
print(df)

page = diagram.getPages().getPage("Page-1")

for index, row in df.iterrows():
    # Extract flowchart elements from the row
    print(row)
    id = row[0]
    description = row["Description"]
    actor = row["Responsible"]
    try:
        type = "Decision" if len(row["Decision"].strip()) else "Process"
    except AttributeError:
        pass

    # Create a FlowchartNode object
    flowchart_node = FlowchartNode(id, None, type, None, description, actor)
    flowchart_nodes.append(flowchart_node)

    # Add the FlowchartNode object to the diagram
    flowchart_node.visioId = diagram.addShape(
        4.5 * len(flowchart_nodes), 5.5, 2, 1, "Rectangle", 0)
    # flowchart_node.visioId = diagram.addShape(
    #     5 * len(flowchart_nodes), 5, 10, 10, flowchart_node.type, 0)
    shape = page.getShapes().getShape(flowchart_node.visioId)
    shape.getText().getValue().add(Txt(f"\n{flowchart_node.description}\n({flowchart_node.actor})"))


# Save Visio file
diagram.save(f"{OUTPUT_DIRECTORY}/output.vsdx", SaveFileFormat.VSDX)
jpype.shutdownJVM()
