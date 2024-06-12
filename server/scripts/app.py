import jpype
import pandas as pd
from flowchart_node import FlowchartNode
import asposediagram
from flask import Flask, request, send_file
from flask_cors import CORS


# Start JVM
if not jpype.isJVMStarted():
    print("Started JVM")
    jpype.startJVM()

from asposediagram.api import *
license = License()
license.setLicense("Aspose.DiagramforPythonviaJava.lic")

# Main script to automate the process of generating Visio flowcharts from notes on a spreadsheet
# TODO: Add Aspose License details
# ? Aspose Docs: https://docs.aspose.com/diagram/java/working-with-visio-shape-data/#use-connection-indexes-to-connect-shapes

app = Flask(__name__)
CORS(app)

OUTPUT_DIRECTORY = "../output"
INPUT_DIRECTORY = "../input"
TEMPLATE_DIRECTORY = "../templates"
PAGE_SPACING_X_FACTOR = 3.3
PAGE_SPACING_Y_FACTOR = 2.5
PAGE_Y_START = 3

@app.route('/generate-flowchart', methods=['POST'])
def generate_flowchart():
    print("Generating flowchart...")
    
    # Read meeting notes spreadsheet file from input directory
    input_data = request.get_json()
    print(input_data)
    if (input_data[0]['Step ID'].lower() == "Step"):
        df = pd.DataFrame(input_data[1:])
    else:
        df = pd.DataFrame(input_data)

    print(df)

    if not jpype.isJVMStarted():
        print("Started JVM")
        jpype.startJVM()

    # Generate Visio file
    template_file = f"{TEMPLATE_DIRECTORY}/Flowchart.vstx"
    diagram = Diagram(template_file)

    flowchart_nodes = {}
    page = diagram.getPages().getPage("Page-1")
    background_page = diagram.getPages().getPage("VBackground-1")
    diagram.getPages().remove(background_page)

    def render_flowchart():
        clear_page(page)

        index = 0
        offset = 0
        nest_level = 1
        max_nest_level = 1

        # stack of nodes to proceed from when proceed popping
        proceed_pop_reference_points = []

        # List all available masters
        for i in range(diagram.getMasters().getCount()):
            print(diagram.getMasters().get(i).getName())
        print()

        while index < len(df):
            print(nest_level)

            row = df.iloc[index]
            print("Analyzing row", index)
            # Extract flowchart elements from the row into FlowchartNode object
            flowchart_node = FlowchartNode.from_spreadsheet_row(row)
            flowchart_node.nest_level = nest_level
            print(flowchart_node.description)

            if flowchart_node.description.split(":")[-1].strip().split(' ')[0] == "Proceed":
                nest_level -= 1
                index += 1
                destination_id = flowchart_node.description.split(
                    ":")[-1].strip().split(' ')[-1]
                print("Proceeding to", destination_id)
                offset -= 1
                print("From:", flowchart_node.id, "Destination:", destination_id)
                if len(flowchart_nodes) >= 1:
                    print("Adding jump from", list(flowchart_nodes.keys())
                          [-1], "to", destination_id)
                    flowchart_nodes[list(flowchart_nodes.keys())
                                    [-1]].add_jump(destination_id)

                continue

            # Add the FlowchartNode object to the diagram
            flowchart_nodes[flowchart_node.id] = flowchart_node

            flowchart_node.visioId = diagram.addShape(
                PAGE_SPACING_X_FACTOR * (index + offset),
                PAGE_Y_START + PAGE_SPACING_Y_FACTOR * nest_level,
                2, 1,
                flowchart_node.type,
                0)

            print("Visio ID:", flowchart_node.visioId)

            render_text(flowchart_node)

            print(flowchart_node.type)

            if len(flowchart_nodes) > 1:
                # If popping out of a nest level, connect previously proceeding nodes to the current node
                print("Proceed Pop Reference Points:",
                      proceed_pop_reference_points)
                print(nest_level)
                print(flowchart_nodes[list(
                    flowchart_nodes.keys())[-2]].description)
                print(flowchart_nodes[list(flowchart_nodes.keys())[-2]].nest_level)

                if len(proceed_pop_reference_points) > 0 and nest_level < flowchart_nodes[list(flowchart_nodes.keys())
                                                                                          [-2]].nest_level:
                    popped_reference_point = proceed_pop_reference_points.pop()
                    print(popped_reference_point.description, "Popped")
                    connect_shapes(popped_reference_point,
                                   flowchart_node, is_sequential=False)
                    flowchart_node.description = popped_reference_point.decisions[0] + ": " + flowchart_node.description
                    render_text(flowchart_node)

                connect_shapes(flowchart_nodes[list(
                    flowchart_nodes.keys())[-2]], flowchart_node, is_sequential=True)

            try:
                if flowchart_node.is_decision():
                    print(flowchart_node.decisions)
                    nest_level += 1
                    max_nest_level += 1
                    proceed_pop_reference_points.append(flowchart_node)
                    print("Proceed Pop Reference Points:",
                          proceed_pop_reference_points)
                    index += 2
                    offset -= 2
                else:
                    index += 1
            except AttributeError as e:
                print("Attribute Error:", e)
                index += 1

            print("Next row to analyze:", index)
            print()

        render_post_connections()

        # ! Not working
        # Todo: fix
        # diagram.getPages().getPage(0).getPageSheet().getPageLayout().getLineAdjustFrom().setValue(LineAdjustFromValue.ALL_LINES);

        # ! Not working
        # Todo: fix
        # flowChartOptions = LayoutOptions()
        # flowChartOptions.setLayoutStyle(LayoutStyle.FLOW_CHART);
        # flowChartOptions.setSpaceShapes(1);
        # flowChartOptions.setEnlargePage(True);
        # flowChartOptions.setDirection(LayoutDirection.LEFT_TO_RIGHT);
        # diagram.layout(flowChartOptions)

        # Save Visio file
        print("Diagram Generated")
        diagram.save(f"output/output.vsdx", SaveFileFormat.VSDX)
        print("Diagram Saved")
        return 0

    def render_post_connections():
        for _, node in flowchart_nodes.items():
            print("Node ID:", node.id, "To IDs:", node.jump_to_ids)
            for to_id in node.jump_to_ids:
                print(node.id, "to", to_id)
                connect_shapes(node, flowchart_nodes[to_id], is_sequential=False)

    def connect_shapes(flowchart_node_1, flowchart_node_2, is_sequential=True):
        if ((not is_sequential) or (flowchart_node_2.nest_level >= flowchart_node_1.nest_level)):
            print("Connecting", flowchart_node_1.id, "to", flowchart_node_2.id)
            connectorShape = Shape()
            connectorId = diagram.addShape(connectorShape, 'Dynamic connector', 0)
            connectorShape.getLayout().getConFixedCode().setValue(ConFixedCodeValue.REROUTE_FREELY)
            connectorLine = connectorShape.getLine()
            connectorLine.getEndArrow().setValue(1)
            connectorLine.getLineWeight().setValue(0.02)
            connectorLine.getLineColor().getUfe().setF("RGB(0,0,0)")

            starting_new_nest_level = (
                flowchart_node_2.nest_level > flowchart_node_1.nest_level)

            page.connectShapesViaConnector(
                flowchart_node_1.visioId,
                ConnectionPointPlace.TOP if starting_new_nest_level else ConnectionPointPlace.RIGHT,
                flowchart_node_2.visioId,
                ConnectionPointPlace.BOTTOM if starting_new_nest_level else ConnectionPointPlace.LEFT,
                connectorId)

    def render_text(flowchart_node):
        print("Rendering text for", flowchart_node.id)
        print(flowchart_node.visioId)
        print(flowchart_node.description)
        # Render FlowchartNode text on page
        shape = page.getShapes().getShape(flowchart_node.visioId)
        shape.getText().getValue().clear()
        shape.getChars().clear()
        shape.getChars().add(Char())
        shape.getChars().add(Char())
        shape.getChars().get(0).getStyle().setValue(StyleValue.BOLD)
        shape.getChars().get(1).getSize().setValue(0.16)

    def clear_page(page):
        shapes = list(page.getShapes())  # Create a copy of the list
        for shape in shapes:
            page.getShapes().remove(shape)

    result = render_flowchart()
    if result == 0:
        return send_file("output/output.vsdx", as_attachment=True)
    else:
        return render_text("Sorry, an error was encountered")
    



# jpype.shutdownJVM()
