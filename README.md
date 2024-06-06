# Process Flowchart Generator

This project is a Python script that automates the process of generating Visio flowcharts from notes on a spreadsheet.

## Project Structure

```
input/
	Meeting_notes.xlsx
LICENSE
output/
	output.vsdx
README.md
scripts/
	flowchart_node.py
	main.py
templates/
	Basic Shapes.vss
	Flowchart.vstx
```

## How it Works

The script reads meeting notes from an Excel spreadsheet file located in the [input](/input) directory. It then processes the meeting notes and generates a Visio file using a template from the [templates](/templates) directory. Currently, the only supported template is [Flowchart.vstx](templates/Flowchart.vstx) The output Visio file is saved in the [output](/output) directory.

The main script is located in [scripts/main.py](/scripts/main.py). It uses the [FlowchartNode](/scripts/flowchart_node.py) class to represent nodes in the flowchart.

## Usage

1. Copy your Meeting notes into [``Meeting_notes.xlsx``](/input/Meeting_notes.xlsx) directory.
2. Run the script with the command `python scripts/main.py`.
3. The generated Visio file will be saved in the [``output``](/output) directory.

## Example Input and Output
### Input:
| Step ID | Description                    | Responsible       | Decision           |
| ------- | ------------------------------ | ----------------- | ------------------ |
| 1       | Initiate New Data Collection   | Data Team Lead    |                    |
| 2       | Condition of Data              |                   | Good/Bad           |
| 3       | Good: Proceed                  |                   |                    |
| 4       | Bad: Recollect Data            | Data Engineer     |                    |
| 5       | Validate Data                  | Data Engineer     |                    |
| 6       | Proceed to 2                   |                   |                    |
| 7       | Send Data                      | Vendor Team       |                    |

### Output:
![Example Flowchart output](Example%20Output.png)

## Dependencies

This project uses the following Python libraries:

- jpype
- pandas
- asposediagram

## License

This project is licensed under the terms of the LICENSE file.

The project also depends on an Aspose Diagram License for full usage. Without it, the generated Visio file is limited to 10 shapes (including connectors)
