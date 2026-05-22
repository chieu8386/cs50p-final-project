import pandas as pd
import glob
import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from openpyxl.utils import get_column_letter


def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        df = read_excel("input_data")
        df = clean_dataframe(df)
        df = filter_columns(df)
        filename = export_excel(df, "output_data")

        categories = df["Category"].unique()
        categories = sorted(categories, key=lambda x: (x == "OTHERS", x))
        thong_bao = f"Xử lý hoàn tất\n\nFile: {os.path.basename(filename)}\nSheet: {', '.join(categories)}"
        messagebox.showinfo("Kết quả", thong_bao)

    except Exception as e:
        messagebox.showerror("Lỗi", f"Chi tiết lỗi: {e}")
    finally:
        root.destroy()


def read_excel(input_folder):
    """Đọc file Excel đầu tiên trong thư mục input"""
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    if not excel_files:
        raise FileNotFoundError(f"Không tìm thấy file Excel trong '{input_folder}'")
    input_file = excel_files[0]
    print(f"Đang xử lý file: {input_file}")
    return pd.read_excel(input_file, header=1)


def clean_dataframe(df):
    """Làm sạch DataFrame: tên cột, kiểu số, lọc hàng, xử lý chuỗi"""
    # Làm sạch tên cột
    df.columns = [str(c).strip() for c in df.columns]

    # Chuyển cột số
    cols_to_fix = [
        'List Price', 'List Price(Unit)', 'Quantity Billed',
        'NetAmount(excTax)', 'NetAmount(Company Curr)', 'NetAmount(Unit)',
        'ActPrice(IncTax)', 'ActPrice(Unit)', 'COGS(Unit)', 'Price Difference'
    ]
    for col in cols_to_fix:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(',', ''),
                errors='coerce'
            )

    # Lọc hàng không cần thiết
    if "CGrp Description" in df.columns:
        df = df[df["CGrp Description"] != "BICV Direct"]
    if "SalesRep Description" in df.columns:
        df = df[df["SalesRep Description"] != "GROUP"]

    # Xử lý tên sản phẩm
    product_col = "Material Description"
    if product_col in df.columns:
        pattern = r"\s+\b(ASA|AP|SGP|VNM/IDN|IDN|US|VNM)$|\s+BH17\b|\s*\(SP\)$"
        df[product_col] = df[product_col].astype(str).str.replace(
            pattern, "", regex=True
        ).str.strip()

    # Thêm cột Category
    if "Material number" in df.columns:
        df["Category"] = df["Material number"].apply(get_category)
    else:
        target_idx = 31 if df.shape[1] > 31 else 0
        df["Category"] = df.iloc[:, target_idx].apply(get_category)

    # Thêm cột Dealer Name
    d_dict = load_mapping()
    try:
        idx_sold_to = df.columns.get_loc("Sold-to")
        dealer_values = df["Sold-to"].apply(lambda name: get_dealer(name, d_dict))
        df.insert(loc=idx_sold_to + 1, column="Dealer Name", value=dealer_values)
    except Exception:
        df["Dealer Name"] = df.iloc[:, 11].apply(lambda name: get_dealer(name, d_dict))

    return df


def filter_columns(df):
    """Lọc và giữ lại các cột cần thiết"""
    required_cols = [
        'M', 'Dealer Name', 'AB', 'AD', 'Category',
        'AH', 'AJ', 'AN', 'AU', 'AW', 'AZ', 'BB', 'BP'
    ]

    def excel_col_to_idx(col_str):
        exp, idx = 0, 0
        for char in reversed(col_str.upper()):
            idx += (ord(char) - ord('A') + 1) * (26 ** exp)
            exp += 1
        return idx - 1

    columns_to_keep = []
    for col_str in required_cols:
        if col_str in ["Category", "Dealer Name"]:
            if col_str in df.columns:
                columns_to_keep.append(col_str)
            continue
        idx = excel_col_to_idx(col_str)
        if idx < len(df.columns):
            columns_to_keep.append(df.columns[idx])

    columns_to_keep = list(dict.fromkeys(columns_to_keep))
    return df[columns_to_keep]


def export_excel(df, output_folder):
    """Xuất DataFrame thành file Excel đa sheet"""
    cols_to_fix = [
        'List Price', 'List Price(Unit)', 'Quantity Billed',
        'NetAmount(excTax)', 'NetAmount(Company Curr)', 'NetAmount(Unit)',
        'ActPrice(IncTax)', 'ActPrice(Unit)', 'COGS(Unit)', 'Price Difference'
    ]

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = os.path.join(output_folder, f"Billing_Reports_{timestamp}.xlsx")

    categories = sorted(
        df["Category"].unique(),
        key=lambda x: (x == "OTHERS", x)
    )

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for cat in categories:
            filtered_df = df[df["Category"] == cat].copy()
            filtered_df = filtered_df.drop(columns=["Category"])
            filtered_df.to_excel(writer, sheet_name=str(cat), index=False)

            worksheet = writer.sheets[str(cat)]

            # Format cột số
            for col_idx, col_name in enumerate(filtered_df.columns):
                if col_name in cols_to_fix:
                    col_letter = get_column_letter(col_idx + 1)
                    for cell in worksheet[col_letter][1:]:
                        cell.number_format = '#,##0'

            # Format cột ngày
            for cell in worksheet['A'][1:]:
                cell.number_format = 'DD/MM/YYYY'
            worksheet.column_dimensions['A'].width = 15

            # Ẩn sheet OTHERS
            if cat == "OTHERS" and len(categories) > 1:
                worksheet.sheet_state = 'hidden'

    return filename

def get_category(model):
    model = str(model).strip().upper()
    if model == "NAN" or not model:
        return "OTHERS"
    
    # 1. Nhóm Laser Màu
    if model.startswith(('84E', '8CE')): return "CL HW"
    
    # 2. Nhóm Laser Đen trắng
    if model.startswith(('84U','8C5','8X5','84UH','84UM','84UF')): return "ML HW"
    
    # 3. Nhóm Máy in phun
    if model.startswith(('84H', '8CH')): return "INK HW"
    
    # 4. Nhóm Mực máy in phun
    
    if model.startswith(('8ZC', '5XX5')):
        return "INK CONS"
    
    # 5. Nhóm Mực/Trống máy Laser
    
    if model.startswith(('84XX', '84GA', '84UX', '84GW',)):
        return "ML CONS"
        
    # 6. Nhóm vật tư Waste Toner box/Belt Unit Color Laser
    if model.startswith(('84GT', '84GC', '84GB', '84GD')): 
        return "CL CONS"
    
    # 7. SCANNERS

    if model.startswith(('5WDE', '5WDD', '5WDF', '5WDC')): return "SCANNER"
    return "OTHERS"

def load_mapping():
    try:
        df_map = pd.read_excel("mapping_dealer.xlsx", sheet_name="Dealer")
        return dict(zip(df_map['Code'].astype(str).str.strip().str.upper(), df_map['Dealer name']))
    except Exception as e:
        print(f"Không tìm thấy file mapping dealer, lỗi {e}")
        return {}
def get_dealer(name, d_dict):
    name_upper = str(name).strip().upper()
    for key, value in d_dict.items():
        if name_upper.endswith(key):
            return value
    return name
      
    
if __name__ == "__main__":
    main()