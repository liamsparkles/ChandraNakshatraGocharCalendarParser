from typing import Tuple

import pandas


def parse_input(filename: str) -> pandas.DataFrame:
    pd_data = pandas.read_excel(filename)
    names = pd_data.iloc[::3, :]
    dates = pd_data.iloc[1::3, :]

    new_dates = pandas.DataFrame()
    new_dates[['Month', 'Day', 'Year', 'Weekday', 'at', 'Time', 'AMPM']] = dates['AllData'].str.split(' ', expand=True)
    new_dates = new_dates.drop(columns='at')
    new_dates = new_dates.map(lambda x: x.strip(","))

    names = names.reset_index().drop(columns="index").rename(columns={"AllData": "Name"})
    new_dates = new_dates.reset_index().drop(columns="index")
    print(names)
    print(new_dates)
    parsed_data = pandas.concat([names, new_dates], axis=1)
    print(parsed_data)
    return parsed_data


def format_output_malay(data: pandas.DataFrame, filename: str) -> None:
    # Date, Time, Name
    formatted_data = pandas.DataFrame()

    # Remake the Date
    formatted_data['Date'] = data['Weekday'] + ", " + data['Month'] + " " + data['Day'] + ', ' + data['Year']

    # Remake the Time
    formatted_data['Time'] = data['Time'] + " " + data["AMPM"]

    # Grab the Name
    formatted_data['Name'] = data['Name']

    formatted_data.to_excel(filename)


def output_simple(data: pandas.DataFrame, filename: str) -> None:
    data.to_excel(filename)


def main() -> None:
    input_folder = "data/input/"
    output_folder = "data/output/"
    input_filename = input_folder + "1900.xlsx"
    output_filename = output_folder + "1900_formatted.xlsx"
    data = parse_input(input_filename)
    format_output_malay(data, output_filename)


if __name__ == "__main__":
    main()
