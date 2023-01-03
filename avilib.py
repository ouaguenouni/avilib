import argparse
from fill import output_completed_xlsx
import os

parser = argparse.ArgumentParser(description='Library helper that does many inventory related tasks.')

parser.add_argument('function', type=str, help='The required functionality.')

parser.add_argument('--input', type=str, nargs='?',help='The path to the xlsx file containing the inventory (the file has to contain an ISBN collumn)')
parser.add_argument('--output', type=str, nargs='?',help='The name of the file in which you want the completed inventory to be stored')
parser.add_argument('--correct', action='store_true', help='A Flag that indicate if the filler will correct.')

args = parser.parse_args()

if args.function == "fill":
    print("Filling the inventory ...")
    if not args.input:
        parser.error("Please provide the path of the inventory to complete..")
    input_name = args.input.split(".")[0]
    if not args.output:
        print(f"Output not provided default output is {input_name}_filled.csv")
        output_path = args.input + "_filled.xlsx"

    if output_path in os.listdir():
        print("Completion already started and will be resumed")
    output_completed_xlsx(args.input, output_path)
else:
    parser.error(f"The function {args.function} is not (yet) implemented :) ")



