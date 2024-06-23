import os
import gpxpy
import pandas as pd
from datetime import datetime


def parse_gpx_file(file_path):
    with open(file_path, 'r') as gpx_file:
        gpx = gpxpy.parse(gpx_file)
        distance = 0.0
        for track in gpx.tracks:
            for segment in track.segments:
                distance += segment.length_3d()  # length_3d returns the length in meters
        distance_km = distance / 1000  # Convert meters to kilometers
        distance_km = round(distance_km)
    return distance_km


def get_date_from_filename(filename):
    date_str = filename[:8]  # Assuming the format starts with YYYYMM
    return datetime.strptime(date_str, '%Y%m%d').date()


def main():
    gpx_folder = './in_gpx/'  # Change this to the path of your GPX files
    data = []

    for filename in os.listdir(gpx_folder):
        if filename.endswith('.gpx'):
            file_path = os.path.join(gpx_folder, filename)
            date = get_date_from_filename(filename)
            mileage = parse_gpx_file(file_path)
            data.append((date, mileage))

    df = pd.DataFrame(data, columns=['Date', 'Mileage'])
    df['YearMonth'] = df['Date'].apply(lambda x: x.strftime('%Y-%m'))
    df = df.sort_values(by='Date', ascending=False)

    # Group by year and month
    monthly_df = df.groupby('YearMonth').agg({'Mileage': 'sum'}).reset_index().sort_values(by='YearMonth', ascending=False)

    # Group by year
    yearly_df = monthly_df.groupby(pd.to_datetime(monthly_df['YearMonth'], format='%Y-%m').dt.year)\
        .agg({'Mileage': 'sum'}).reset_index()
    yearly_df.columns = ['Year', 'Mileage']
    yearly_df = yearly_df.sort_values(by='Year', ascending=False)

    # Save to Excel
    with pd.ExcelWriter('mileage_by_year_month.xlsx') as writer:
        df.to_excel(writer, sheet_name='Daily Mileage', index=False)
        monthly_df.to_excel(writer, sheet_name='Monthly Mileage', index=False)
        yearly_df.to_excel(writer, sheet_name='Yearly Mileage', index=False)

    print("Excel file 'mileage_by_year_month.xlsx' created successfully.")


if __name__ == '__main__':
    main()

