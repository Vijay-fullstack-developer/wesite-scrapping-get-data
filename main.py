from bs4 import BeautifulSoup
import openpyxl, requests
import pandas as pd

excel_file_path = 'DAP_Products_Input_File.xlsx'
sheet_name = 'Input File'
try:
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

    # Create the 'url' column
    df['url'] = 'https://www.lenovo.com/' + df['Country'].astype(str).str.lower() + '/' + df['Locales'].astype(str) + '/p/' + df['SKU'].astype(str)

except ValueError:
    print(f"Worksheet named '{sheet_name}' not found.")
    # You can add further handling or simply exit the program gracefully

# Create an Excel writer
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Data Fetch'
sheet.append(["SKU", "Country", "Locales", "Title", "Description","CTA URL","Link"])

# Iterate over each URL in the DataFrame
for index, row in df.iterrows():
    url = row['url']
    url_inside = ''  # Initialize url_inside here to handle cases where the condition is not met
    try:
        resource = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'})
        resource.raise_for_status()
        soup = BeautifulSoup(resource.text, 'html.parser')

        # Replace condition_true with your actual condition
        condition_true = soup.find('ul', class_="configuratorItem-mtmTable") is not None

        if condition_true:
            table_data = soup.find('ul', class_="configuratorItem-mtmTable").find_all("li")
        else:
            table_data = soup.find('div', class_="system_specs_container").find_all("li")

            New_condition = soup.find('a', class_="view-all-models") is not None

            url_inside = None  # Initialize url_inside

            if New_condition:
                table_data2 = soup.find('a', class_="view-all-models")
            else:
                table_data2 = soup.find('a', class_="clickHereLinkText")

            if table_data2:
                href_value = table_data2.get('href')
                path_without_spaces = '-'.join(href_value.split('/'))
                href_value = href_value.replace(' ', '-').replace('/' + path_without_spaces + '/',
                                                                  '/' + path_without_spaces + '/')
                url_inside = "https://www.lenovo.com/" + str(href_value)


        # Add SKU and Country data to each row
        sku = row['SKU']
        country = row['Country']
        Locales = row['Locales']

        # Iterate over each item in the table_data
        for data in table_data:
            # Extract relevant information
            if condition_true:
                name_tag = data.find('h4', class_="configuratorItem-mtmTable-title").text.strip()
                sub_head_tag = data.find('p').text.strip()
            else:
                name_tag = data.find('div', class_="title").text
                sub_head_tag = data.find('p').text.strip()


            sheet.append([sku, country, Locales, name_tag, sub_head_tag,url, url_inside])



    except Exception as e:
        print(f"Error for URL {url}: {e}")

excel.save("get-data.xlsx")

get_data = 'get-data.xlsx'
sheet_name = 'Data Fetch'
df1 = pd.read_excel(get_data, sheet_name=sheet_name)

get_data2 = 'Transulate_Final_input.xlsx'
sheet_name2 = 'Attribute_Translations'
df2 = pd.read_excel(get_data2, sheet_name=sheet_name2)

df1['Title Name_lower'] = df1['Title'].str.lower()
df2['Otherlanguage_lower'] = df2['Otherlanguage'].str.lower()

df1['Execution Result'] = df1['Title Name_lower'].isin(df2['Otherlanguage_lower']).tolist()

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "fetch data"

sheet.append(["SKU", "Country", "Locales", "Description", "Result","CTA URL","Link"])

for index, row in df1.iterrows():
    title_name_lower = row['Title Name_lower']
    if title_name_lower in df2['Otherlanguage_lower'].values:
        english_row = df2[df2['Otherlanguage_lower'] == title_name_lower]['English'].values[0]
        sheet.append([row['SKU'],None,None , row['Locales'], row['Description'], str(english_row),row['CTA URL'], row['Link'],row['Country']])
    else:
        title_name = row['Title']

        sheet.append([row['SKU'],None,None , row['Locales'], row['Description'],title_name, row['CTA URL'], row['Link'],row['Country']])



# Save the Excel file
excel.save('result.xlsx')













