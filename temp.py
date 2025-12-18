import pandas as pd
import os

FILE_PATH = 'StudentData.xlsx'

def save_data(df):
    """Lưu dữ liệu vào file Excel."""
    try:
        df.to_excel(FILE_PATH, index=False)
        print("💾 Hệ thống: Đã tự động lưu thay đổi vào file.")
    except PermissionError:
        print("⚠️ Cảnh báo: Không thể lưu vì file Excel đang mở. Hãy đóng file và thử lại!")

def analysis(TX, GK, CK):
    """Tính toán điểm trung bình, GPA và xếp loại học lực."""
    TB = round(TX * 0.1 + GK * 0.3 + CK * 0.6, 2)
    if TB >= 8.5: 
        GPA, HL = 4.0, 'Xuất sắc'
    elif TB >= 8.0: 
        GPA, HL = 3.5, 'Khá giỏi'
    elif TB >= 7.0: 
        GPA, HL = 3.0, 'Khá'
    elif TB >= 6.5: 
        GPA, HL = 2.5, 'Trung bình khá'
    elif TB >= 5.5: 
        GPA, HL = 2.0, 'Trung bình'
    elif TB >= 5.0:
        GPA, HL = 1.5, 'Trung bình yếu'
    elif TB >= 4.0:
        GPA, HL = 1.0, 'Yếu'
    else: 
        GPA, HL = 0.0, 'Kém'
    return (TB, GPA, HL)

def add_student(df):
    """Thêm học sinh mới vào danh sách."""
    print("\n--- THÊM HỌC SINH ---")
    try:
        student_id = int(input("Nhập ID học sinh: "))
        if student_id in df['ID'].values:
            print(f"❌ Lỗi: ID {student_id} đã tồn tại.")
            return df
            
        name = input("Nhập tên học sinh: ")
        gender = input('Nhập giới tính: ')
        TX = float(input("Nhập điểm thường xuyên: "))
        GK = float(input("Nhập điểm giữa kỳ: "))
        CK = float(input("Nhập điểm cuối kỳ: "))
        DRL = float(input('Nhập điểm rèn luyện: '))
        
        TB, GPA, HL = analysis(TX, GK, CK)
        
        new_row = pd.DataFrame([{
            'ID': student_id, 'Name': name, 'Giới tính': gender, 
            'TX': TX, 'GK': GK, 'CK': CK, 'DRL': DRL, 
            'Tổng': TB, 'GPA': GPA, 'Học lực': HL
        }])
        
        df = pd.concat([df, new_row], ignore_index=True)
        print(f"✅ Đã thêm học sinh {name} thành công.")
        save_data(df)
        return df
    except ValueError:
        print("❌ Lỗi: Sai định dạng số. Vui lòng thử lại.")
        return df

def change_score(df, student_id, score_type):
    """Cập nhật điểm thành phần và tính toán lại toàn bộ kết quả."""
    idx = df[df['ID'] == student_id].index
    if not idx.empty:
        try:
            new_val = float(input(f"Nhập điểm {score_type} mới: "))
            df.loc[idx, score_type] = new_val
            
            # Lấy các điểm hiện tại để tính toán lại
            tx = df.loc[idx, 'TX'].values[0]
            gk = df.loc[idx, 'GK'].values[0]
            ck = df.loc[idx, 'CK'].values[0]
            
            tb, gpa, hl = analysis(tx, gk, ck)
            
            # Cập nhật kết quả mới vào DataFrame
            df.loc[idx, 'Tổng'] = tb
            df.loc[idx, 'GPA'] = gpa
            df.loc[idx, 'Học lực'] = hl
            
            print(f"✅ Đã cập nhật điểm {score_type} và tính lại xếp loại.")
            save_data(df)
        except ValueError:
            print("❌ Lỗi: Điểm nhập vào phải là số.")
    return df

def search_by_id(df):
    """Tìm kiếm học sinh và cung cấp menu sửa điểm."""
    print("\n--- TÌM KIẾM THEO ID ---")
    if df.empty:
        print("⚠ Danh sách hiện đang trống.")
        return df # Luôn trả về df để tránh lỗi NoneType
        
    try:
        search_id = int(input("Nhập ID cần tìm: "))
        result = df[df['ID'] == search_id]
        
        if not result.empty:
            print("\n" + result.to_string(index=False))
            print("\nBạn có muốn sửa điểm cho học sinh này?")
            print("1. Có | 2. Không")
            choice = input("Chọn: ")
            
            if choice == '1':
                while True:
                    print("\n--- MENU SỬA ĐIỂM ---")
                    print("1. Thường xuyên | 2. Giữa kỳ | 3. Cuối kỳ | 4. Rèn luyện | 5. Quay lại")
                    sub_choice = input("Chọn (1-5): ")
                    
                    if sub_choice == '5': break
                    
                    mapping = {'1': 'TX', '2': 'GK', '3': 'CK', '4': 'DRL'}
                    score_type = mapping.get(sub_choice)
                    
                    if score_type:
                        df = change_score(df, search_id, score_type)
                    else:
                        print("⚠ Lựa chọn không hợp lệ.")
        else:
            print(f"❌ Không tìm thấy học sinh có ID: {search_id}")
    except ValueError:
        print("❌ Lỗi: ID phải là một số nguyên.")
    
    return df # Đảm bảo luôn trả về df cho hàm main

def display_all(df):
    """Hiển thị toàn bộ danh sách."""
    print("\n--- DANH SÁCH HỌC SINH ---")
    if df.empty:
        print("⚠ Danh sách hồ sơ rỗng.")
    else:
        print(df.sort_values(by='ID').to_string(index=False))

def main():
    # Khởi tạo dữ liệu
    if os.path.exists(FILE_PATH):
        try:
            df = pd.read_excel(FILE_PATH)
        except:
            df = pd.DataFrame(columns=['ID', 'Name','Giới tính', 'TX', 'GK', 'CK', 'DRL', 'Tổng', 'GPA', 'Học lực'])
    else:
        df = pd.DataFrame(columns=['ID', 'Name','Giới tính', 'TX', 'GK', 'CK', 'DRL', 'Tổng', 'GPA', 'Học lực'])

    while True:
        print("\n" + "="*30)
        print("      QUẢN LÝ LỚP HỌC")
        print("="*30)
        print("1. Thêm học sinh")
        print("2. Tìm kiếm & Sửa điểm")
        print("3. Hiển thị danh sách")
        print("4. Lưu & Thoát")
        
        choice = input("Chọn (1-4): ")
        
        if choice == '1':
            df = add_student(df)
        elif choice == '2':
            df = search_by_id(df) # Gán lại df để không bị NoneType
        elif choice == '3':
            display_all(df)
        elif choice == '4':
            save_data(df)
            print("👋 Đã thoát chương trình. Tạm biệt!")
            break
        else:
            print("⚠ Lựa chọn không hợp lệ.")

if __name__ == "__main__":
    main()