
import pandas as pd
import os

FILE_PATH = 'StudentData.xlsx'
def save_data(df):
    try:
        df.to_excel(FILE_PATH, index=False)
        print("💾 Đã tự động sao lưu dữ liệu.")
    except PermissionError:
        print("⚠️ Cảnh báo: Không thể tự động lưu vì file Excel đang mở!")

def analysis(TX,GK,CK):
    TB = round( TX * 0.1 + GK * 0.3 + CK * 0.6,2)
    if TB >= 8.5: 
        GPA = 4.0
        HL = 'Xuất sắc'
    elif TB >= 8.0: 
        GPA = 3.5
        HL = 'Khá giỏi'
    elif TB >= 7.0: 
        GPA = 3.0
        HL = "Khá"
    elif TB >= 6.5: 
        GPA = 2.5
        HL = 'Trung bình khá'
    elif TB >= 5.5: 
        GPA = 2.0
        HL = 'Trung bình'
    elif TB >= 5.0:
        GPA = 1.5
        HL = 'Trung bình yếu'
    elif TB >= 4.0:
        GPA = 1.0
        HL = 'Yếu'
    else: 
        GPA = 0.0
        HL = 'Kém'
    return (TB, GPA, HL)

def add_student(df):
    """Nhận DataFrame và trả về DataFrame mới sau khi thêm học sinh."""
    print("\n--- Thêm học sinh ---")
    try:
        student_id = int(input("Nhập ID học sinh: "))
        
        # Kiểm tra ID trùng lặp bằng cách dùng .isin() của Pandas
        if student_id in df['ID'].values:
            print(f"❌ Lỗi: ID {student_id} đã tồn tại.")
            return df
            
        name = input("Nhập tên học sinh: ")
        SE = input('Nhập giới tính: ')
        TX = float(input("Nhập điểm thường xuyên: "))
        GK = float(input("Nhập điểm giữa kỳ: "))
        CK = float(input("Nhập điểm cuối kỳ: "))
        DRL= float(input('Nhập điểm rèn luyện: '))
        TB, GPA, HL = analysis(TX,GK,CK)
        # Tạo DataFrame mới từ dictionary
        new_row = pd.DataFrame([{'ID': student_id, 'Name': name, 'Giới tính': SE, 'TX': TX, 'GK': GK, 'CK': CK, 'DRL': DRL, 'Tổng': TB, 'GPA': GPA, 'Học lực': HL}])
        
        # Nối vào DataFrame cũ
        df = pd.concat([df, new_row], ignore_index=True)
        print(f"✅ Đã thêm học sinh {name} thành công.")
        return df
        
    except ValueError:
        print("❌ Lỗi: ID phải là số nguyên và Điểm số phải là số thập phân.")
        return df

def change_score(df, id, score_type):
    idx = df[df['ID'] == id].index
    if not idx.empty:
        try:
            new_val = float(input(f"Nhập điểm {score_type} mới: "))
            df.loc[idx, score_type] = new_val
            
            # Tính toán lại các cột dựa trên điểm mới
            TX = df.loc[idx, 'TX'].values[0]
            GK = df.loc[idx, 'GK'].values[0]
            CK = df.loc[idx, 'CK'].values[0]
            
            TB, GPA, HL = analysis(TX, GK, CK)
            
            df.loc[idx, 'Tổng'] = TB
            df.loc[idx, 'GPA'] = GPA
            df.loc[idx, 'Học lực'] = HL
            
            save_data(df)
            print("✅ Đã cập nhật và tính toán lại xếp loại.")
        except ValueError:
            print("❌ Lỗi: Điểm nhập vào phải là số.")
    return df

def search_by_id(df):
    print("\n--- Tìm kiếm theo ID ---")

    try:
        search_id = int(input("Nhập ID cần tìm: "))
        # Lọc dữ liệu
        result = df[df['ID'] == search_id]
        if not result.empty:
            print("\n--- Kết quả tìm thấy ---")
            print(result.to_string(index=False))
            print("Có muốn sửa điểm sinh viên này hay không??!! ")
            print("""1. Có
2. Không""")
            check = int(input('chọn 1/2: '))
            if check == 1:
                while True:
                    print("""Chọn điểm cần sửa:
1. Điểm thường xuyên
2. Điểm giữa kỳ
3. Điểm cuối kỳ
4. Điểm rèn luyện
5. Thoát""")
                    chon = int(input("Chọn từ 1 - 5: "))
                    match chon:
                        case 1: type = 'TX'
                        case 2: type = 'GK'
                        case 3: type = 'CK'
                        case 4: type = 'DRL'
                        case 5: break
                    df = change_score(df, search_id, type) 
        else:
            print(f"❌ Không tìm thấy học sinh có ID: {search_id}")
    except ValueError:
        print("❌ Lỗi: ID nhập vào phải là số.")
        return df

def display_all_scores(df):
    print("\n--- Danh sách học sinh ---")
    if df.empty:
        print("⚠ Danh sách hồ sơ rỗng.")
    else:
        # Sắp xếp theo ID cho dễ nhìn trước khi in
        print(df.sort_values(by='ID').to_string(index=False))

def main():
    # Khởi tạo hoặc đọc dữ liệu
    if os.path.exists(FILE_PATH):
        try:
            df = pd.read_excel(FILE_PATH)
        except Exception as e:
            print(f"Lỗi đọc file: {e}. Đang tạo mới...")
            df = pd.DataFrame(columns=['ID', 'Name','Giới tính', 'TX', 'GK', 'CK', 'DRL', 'Tổng', 'GPA', 'Học lực'])
    else:
        df = pd.DataFrame(columns=['ID', 'Name','Giới tính', 'TX', 'GK', 'CK', 'DRL', 'Tổng', 'GPA', 'Học lực'])

    while True:
        print("\n" + "="*25)
        print("  CLASSROOM MANAGER")
        print("="*25)
        print("1. Thêm học sinh")
        print("2. Tìm kiếm ID")
        print("3. Hiển thị danh sách")
        print("4. Lưu & Thoát")
        
        choice = input("Chọn (1-4): ")
        
        if choice == '1':
            df = add_student(df)
        elif choice == '2':
            df = search_by_id(df)
        elif choice == '3':
            display_all_scores(df)
        elif choice == '4':
            try:
                # Cố gắng ghi file
                df.to_excel(FILE_PATH, index=False)
                print("✅ Dữ liệu đã được lưu thành công vào StudentData.xlsx")
                break
            except PermissionError:
                print("❌ Lỗi: Không thể lưu! Hãy đóng file Excel 'StudentData.xlsx' đang mở và thử lại.")
            except Exception as e:
                print(f"❌ Lỗi không xác định khi lưu: {e}")
        else:
            print("⚠ Lựa chọn không hợp lệ, vui lòng nhập lại.")

if __name__ == "__main__":
    main()