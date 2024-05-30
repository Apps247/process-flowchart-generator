# Main script to automate the process of generating Visio flowcharts from notes on a spreadsheet
# TODO: Add Aspose License details
# ? Aspose Docs: https://docs.aspose.com/diagram/java/working-with-visio-shape-data/#use-connection-indexes-to-connect-shapes


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

PAGE_SPACING_X_FACTOR = 2.5
PAGE_SPACING_Y_FACTOR = 2.5

# Read meeting notes spreadsheet file from input directory
input_file = f"{INPUT_DIRECTORY}/Meeting_notes.xlsx"
df = pd.read_excel(input_file)

# Generate Visio file
template_file = f"{TEMPLATE_DIRECTORY}/Basic Shapes.vss"
diagram = Diagram(template_file)

flowchart_nodes = []
page = diagram.getPages().getPage("Page-1")


def render_flowchart():
    index = 0
    offset = 0
    nest_level = 1
    max_nest_level = 1

    in_alternate_path = False

    print(df)

    while index < len(df):
        row = df.iloc[index]
        print(index)
        # Extract flowchart elements from the row into FlowchartNode object
        flowchart_node = FlowchartNode.from_spreadsheet_row(row)

        print(flowchart_node.description)

        # if flowchart_node.description.split(":")[-1].strip().split(' ')[0] == "Proceed":
        if flowchart_node.description.split(":")[-1].strip().split(' ')[0] == "Proceed":
            nest_level -= 1
            index += 1
            print("Proceed-popping")
            print("Proceed to", flowchart_node.description.split(
                ":")[-1].strip().split(' ')[-1])
            offset -= 1
            continue

        flowchart_nodes.append(flowchart_node)

        # Add the FlowchartNode object to the diagram

        flowchart_node.visioId = diagram.addShape(
            PAGE_SPACING_X_FACTOR * (index + offset),
            PAGE_SPACING_Y_FACTOR * nest_level, 2, 1, "Rectangle", 0)

        render_text(flowchart_node)

        print(flowchart_node.type)
        try:
            if flowchart_node.is_decision():
                print(flowchart_node.decisions)
                nest_level += 1
                max_nest_level += 1
                index += 2
                offset -= 2
            else:
                index += 1
        except AttributeError as e:
            print("Attribute Error:", e)
            index += 1

        print("End", index)

    page.getPageSheet().getPageProps().getPageHeight().setValue(
        PAGE_SPACING_Y_FACTOR * max_nest_level)
    page.getPageSheet().getPageProps().getPageWidth().setValue(
        PAGE_SPACING_X_FACTOR * (len(flowchart_nodes) + offset))

    # Save Visio file
    diagram.save(f"{OUTPUT_DIRECTORY}/output.vsdx", SaveFileFormat.VSDX)


def render_text(flowchart_node):
    # Render FlowchartNode text on page
    shape = page.getShapes().getShape(flowchart_node.visioId)
    text = f"{flowchart_node.description}\n({flowchart_node.actor})".replace(
        "(nan)", "")
    shape.getText().getValue().add(Txt(text))


if __name__ == "__main__":
    render_flowchart()

jpype.shutdownJVM()
