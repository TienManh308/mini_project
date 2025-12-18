import pandas as pd
import os # Dùng để kiểm tra file tồn tại

FILE_PATH = 'StudentData.xlsx'

def add_student(df):   
    """Nhận DataFrame và trả về DataFrame mới sau khi thêm học sinh."""
    print("\n--- Thêm học sinh ---")
    
    try:
        student_id = int(input("Nhập ID học sinh: "))
        name = input("Nhập tên học sinh: ")
        score = float(input("Nhập điểm số: "))
        
        # Kiểm tra ID trùng lặp (Tùy chọn)
        if student_id in df['ID'].values:
            print(f"❌ Lỗi: ID {student_id} đã tồn tại.")
            return df
            
        new_student = {'ID': student_id, 'Name': name, 'Score': score}
        
        # Tạo DataFrame tạm thời cho học sinh mới và nối (concat) vào DataFrame chính
        new_row_df = pd.DataFrame([new_student])
        new_df = pd.concat([df, new_row_df], ignore_index=True)
        
        print(f"✅ Đã thêm học sinh {name} (ID: {student_id}).")
        
        return new_df # Trả về DataFrame đã cập nhật
        
    except ValueError:
        print("❌ Lỗi: ID và Điểm số phải là số.")
        return df # Trả lại df ban đầu nếu có lỗi

def search_by_id(df):
    #Tìm kiếm và hiển thị hồ sơ theo ID 
    print("\n--- Search by ID---")
    if df.empty:
        print("No student found")
        return
        
    try:
        search_id = int(input("ID Input: "))
        
        # Lọc DataFrame
        result = df[df['ID'] == search_id]
        
        if not result.empty:
            print("--- Hồ sơ tìm thấy ---")
            # In dữ liệu của dòng đầu tiên tìm được
            student_data = result.iloc[0] 
            print(f"ID: {student_data['ID']}")
            print(f"Tên: {student_data['Name']}")
            print(f"Điểm số: {student_data['Score']}")
        else:
            print(f"❌ Không tìm thấy học sinh có ID: {search_id}")
            
    except ValueError:
        print("❌ Lỗi: ID phải là số.")
    except KeyError:
        # Xảy ra nếu tên cột không khớp (vd: không có cột 'ID')
        print("❌ Lỗi: Không tìm thấy cột 'ID' trong file dữ liệu.")

def display_all_scores(df):
    """Hiển thị tất cả điểm số."""
    print("\n--- Tất cả điểm số ---")
    if df.empty:
        print("Danh sách hồ sơ rỗng.")
        return

    # In DataFrame bằng hàm to_string để căn chỉnh tốt hơn trong terminal
    print(df.to_string(index=False)) 



def main():
    try:
        # Nếu file tồn tại, đọc file
        df = pd.read_excel(FILE_PATH)
        

    except FileNotFoundError:
        # Nếu không có file, tạo DataFrame rỗng
        df = pd.DataFrame(columns=['ID', 'Name', 'Score']) 
    '''except Exception as e:
        print(f"Lỗi khi đọc file Excel: {e}")
        return # Thoát nếu không thể đọc được file '''

    while True:
        print("\n=== Classroom Data Manager ===")
        print("1. Thêm học sinh mới")
        print("2. Tìm kiếm theo ID")
        print("3. Hiển thị tất cả điểm số")
        print("4. Thoát & Lưu dữ liệu")
        
        choice = input("Nhập lựa chọn của bạn (1-4): ")
        
        if choice == '1':
            df = add_student(df) # Cập nhật df bằng DataFrame mới
        elif choice == '2':
            search_by_id(df) 
        elif choice == '3':
            display_all_scores(df)
        elif choice == '4':
            print("--- Đang lưu dữ liệu ---")
            try:
                # Ghi DataFrame đã cập nhật trở lại file Excel
                df.to_excel(FILE_PATH, index=False)
                print("Saving Successfully")
            except Exception as e:
                print("Error")
                
            print("QUITTING...")
            break
        else:
            print("Please choose valid numbers")


if __name__ == "__main__":
    main() # chay ham main 