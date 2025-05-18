**Mô tả:**

Hệ thống điểm danh bằng nhận diện khuôn mặt sử dụng deep learning và OpenCV. Hệ thống có thể:

Nhận diện khuôn mặt từ camera

Ghi lại thời gian vào/ra

Xuất báo cáo điểm danh dưới dạng file Excel kèm hình ảnh


**Yêu cầu hệ thống:**
Phần cứng
Webcam

RAM tối thiểu 4GB (khuyến nghị 8GB trở lên)

GPU (khuyến nghị để tăng tốc độ nhận diện)

Phần mềm
Python 3.7+

Các thư viện cần thiết


**Cài đặt thư viện:**
Cài đặt các thư viện cần thiết bằng pip:

bash

pip install opencv-python==4.5.5.64

pip install numpy==1.21.6

pip install pandas==1.3.5

pip install tensorflow==2.8.0

pip install keras==2.8.0

pip install pillow==9.2.0

pip install openpyxl==3.0.10

pip install tk==0.1.0

Hoặc cài đặt tất cả bằng file requirements.txt:

bash

pip install -r requirements.txt


**Cách chạy chương trình:**
Clone repository này

Đảm bảo bạn đã cài đặt tất cả các thư viện cần thiết

Tải model đã train sẵn và đặt trong thư mục gốc (hoặc thay đổi đường dẫn trong code dòng số 21)

Chạy file main:

bash
python main.py

**Giao diện chương trình:**
Chương trình có giao diện đơn giản với 2 nút chính:

Điểm danh: Bắt đầu quá trình nhận diện khuôn mặt

Thoát: Đóng chương trình


**Đầu ra:**
Chương trình sẽ tạo file Excel trong thư mục được chỉ định với các thông tin:

Tên người điểm danh

Thời gian vào

Thời gian ra

Trạng thái đủ tiết (Có/Không)

Hình ảnh khuôn mặt


**Lưu ý:**
Đảm bảo đường dẫn đến model (pp1.h5) là chính xác

File Excel sẽ được lưu tự động khi đóng chương trình

Hình ảnh tạm sẽ tự động xóa sau khi lưu vào file Excel

**Giấy phép:**
Dự án này được phân phối theo giấy phép MIT. Xem file LICENSE để biết thêm chi tiết.
