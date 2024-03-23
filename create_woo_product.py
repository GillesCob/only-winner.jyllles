from woocommerce import API
import openpyxl


scrapping_month="Mars"

excel_data_file_path = f'/Users/gillescobigo/Documents/Gilles/Dev/Only Winners/DATAS/2024/{scrapping_month}/EXCEL/DATAS.xlsx'
try:
    classeur_excel_month_exist = openpyxl.load_workbook(excel_data_file_path)
    feuille_excel_month_exist = classeur_excel_month_exist.active
except FileNotFoundError:
    classeur_exist = None

classeur_bdd_WP = openpyxl.load_workbook(excel_data_file_path)
WP_data_sheet = classeur_bdd_WP["Data for WP"]

#------------------Récupération des données dans l'Excel BDD INITIALE-----------------------------------#
ID_list = [cell.value for cell in WP_data_sheet['A'] if cell.value is not None]
Type_list = [cell.value for cell in WP_data_sheet['B'] if cell.value is not None]
SKU_list = [cell.value for cell in WP_data_sheet['C'] if cell.value is not None]
Name_list = [cell.value for cell in WP_data_sheet['D'] if cell.value is not None]
Published_list = [cell.value for cell in WP_data_sheet['E'] if cell.value is not None]
Is_featured_list = [cell.value for cell in WP_data_sheet['F'] if cell.value is not None]
Visibility_in_catalog_list = [cell.value for cell in WP_data_sheet['G'] if cell.value is not None]
Tax_status_list = [cell.value for cell in WP_data_sheet['L'] if cell.value is not None]
In_stock_list = [cell.value for cell in WP_data_sheet['N'] if cell.value is not None]
Backorders_allowed_list = [cell.value for cell in WP_data_sheet['P'] if cell.value is not None]
Sold_individually_list = [cell.value for cell in WP_data_sheet['Q'] if cell.value is not None]
Regular_price_list = [cell.value for cell in WP_data_sheet['Y'] if cell.value is not None]
Categories_list = [cell.value for cell in WP_data_sheet['Z'] if cell.value is not None]
Tags_list = [cell.value for cell in WP_data_sheet['AA'] if cell.value is not None]
Images_list = [cell.value for cell in WP_data_sheet['AC'] if cell.value is not None]
Position_list = [cell.value for cell in WP_data_sheet['AL'] if cell.value is not None]
Attribute_1_name_list = [cell.value for cell in WP_data_sheet['AM'] if cell.value is not None]
Attribute_1_values_list = [cell.value for cell in WP_data_sheet['AN'] if cell.value is not None]
Attribute_1_visible_list = [cell.value for cell in WP_data_sheet['AO'] if cell.value is not None]
Attribute_2_name_list = [cell.value for cell in WP_data_sheet['AQ'] if cell.value is not None]
Attribute_2_values_list = [cell.value for cell in WP_data_sheet['AR'] if cell.value is not None]
Attribute_2_visible_list = [cell.value for cell in WP_data_sheet['AS'] if cell.value is not None]
Attribute_3_name_list = [cell.value for cell in WP_data_sheet['AU'] if cell.value is not None]
Attribute_3_values_list = [cell.value for cell in WP_data_sheet['AV'] if cell.value is not None]
Attribute_3_visible_list = [cell.value for cell in WP_data_sheet['AW'] if cell.value is not None]
Attribute_4_name_list = [cell.value for cell in WP_data_sheet['AY'] if cell.value is not None]
Attribute_4_values_list = [cell.value for cell in WP_data_sheet['AZ'] if cell.value is not None]
Attribute_4_visible_list = [cell.value for cell in WP_data_sheet['BA'] if cell.value is not None]
Attribute_5_name_list = [cell.value for cell in WP_data_sheet['BC'] if cell.value is not None]
Attribute_5_values_list = [cell.value for cell in WP_data_sheet['BD'] if cell.value is not None]
Attribute_5_visible_list = [cell.value for cell in WP_data_sheet['BE'] if cell.value is not None]

for index, id in enumerate(ID_list) :
    id = id
    type = Type_list[index]
    sku = SKU_list[index]
    name = Name_list[index]
    published = Published_list[index]
    is_feature = Is_featured_list[index]
    visibility_in_catalog = Visibility_in_catalog_list[index]
    tax_status = Tax_status_list[index]
    in_stock = In_stock_list[index]
    backorders_allowed = Backorders_allowed_list[index]
    sold_individualy = Sold_individually_list[index]
    regular_price = Regular_price_list[index]
    categories = Categories_list[index]
    tags = Tags_list[index]
    images = Images_list[index]
    position = Position_list[index]
    att_1_name = Attribute_1_name_list[index]
    att_1_value = Attribute_1_values_list[index]
    att_1_visible = Attribute_1_visible_list[index]
    att_2_name = Attribute_2_name_list[index]
    att_2_value = Attribute_2_values_list[index]
    att_2_visible = Attribute_2_visible_list[index]
    att_3_name = Attribute_3_name_list[index]
    att_3_value = Attribute_3_values_list[index]
    att_3_visible = Attribute_3_visible_list[index]
    att_4_name = Attribute_4_name_list[index]
    att_4_value = Attribute_4_values_list[index]
    att_4_visible = Attribute_4_visible_list[index]
    att_5_name = Attribute_5_name_list[index]
    att_5_value = Attribute_5_values_list[index]
    att_5_visible = Attribute_5_visible_list[index]



wcapi = API(
    url="https://www.only-winners.jyllles.com",
    consumer_key="ck_48d46e8dfb551842f07558b27217e0689d1ab517",
    consumer_secret="cs_49754a432678311fba793574a04ba1470177e0f3",
    wp_api=True,
    version="wc/v3"
)

data = {
    "name": name,
    "type": type,
    "regular_price": regular_price,
    "description": "",
    "short_description": "",
    "categories": [
        {
            "id": 1,
            "name": att_1_value
            
        },
    ],
    "images": [
        {
            "src": "https://www.only-winners.jyllles.com/wp-content/uploads/2024/03/19-Judo-Grand-Slam-Tashkent-Womens-Half-Middleweight-63-kg-March-2nd-2024.png"
        },
    ]
}

print(wcapi.post("products", data).json())