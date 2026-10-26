# CNTT1-nh-p-m-n-c-ng-ngh-th-ng-tin-Session12_Ex03
# Phần 1.Phân tích : mapping 4 tham số của VLOOKUP vào mô hình "Nhân viên kho tra hồ sơ"
 ## cú pháp chuẩn của hàm VLOOKUP : VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup]
| Tham số VLOOKUP | Tên gọi kỹ thuật | Tương đương trong mô hinh "Nhân Viên Kho | Ý nghĩa thực tế FastMove |
| --- | --- | --- | --- | 
| look_value | giá trị cần tra cứu | Tờ giấy ghi mã vận đơn mà nhân viên kho đang cầm trên tay | Mã vận đơn duy nhất ( ví dụ : `FM-10023` ) tại ô A2 của sheet danh sách đơn hàng |
| table_array | bảng chứa dữ liệu tra cứu | tủ tài liệu/Bảng giá cước lưu trữ ở căn phòng khác ( sheet `BangGiaCuoc` ). | Vùng dữ liệu từ cột mã đến cột giá cước , sử dụng tham chiếu tuyệt đối  `BangGiaCuoc!$A$2:$C$100`
|col_index_num | thứ tự cột lấy dữ liệu | số thứ tự của ngăn kéo trong tủ tài liệu chứa thông giá cước ( cột thứ N ) | Chỉ số cột chứa thông tin giá cước |
|range_lookup | phương thức tìm kiếm | quy tắc đối soát: Tìm chính xác 100% | bắt buộc chọn `FALSE` để đảm bảo mã vận đơn phải trùng hớp hoàn toàn mới lấy giá cước |
# Phần 2. Triển khai : giải pháp công nghệ cho FastMove
 ## Word 
  cấu trúc dữ liệu: 
  1. Quy trình tiếp nhận và phân loại hàng hóa ( heading 1 )

      1.1 Quét mã vận đơn và đối soát cước tự động ( heading 2 )

      1.2 kiểm tra trọng lượng và kích thước quy đổi ( heading 2 )

  2. Hướng dẫn shippe xuất giao và giao hàng ( heading 1 )
 ## Excel - nối 2 sheet khác nhau
  Đảm bảo không vi phạm 5 quy tắc vàng :
   1. Không merge cell: Tuyệt đối không gộp ô ở vùng dữ liệu để công thức kéo không bị lỗi `#VALUE` hoặc chệch hàng
   2. Dòng header duy nhất : Dòng 1 chứa tiêu đề cột ( mã vận đơn , khối lượng , vùng giao , giá cước )
   3. dữ liệu liên tục : không chứa dòng trống/cột trống xen kẽ giữa các bản ghi.
   4. Đúng dữ liệu ( database)
   5. không chứa dòng tổng cộng ở giữa dữ liệu
* Công thức VLOOKUP liên kết 2 sheet cho 500 dòng

  -sheet 1 ( `danhsachdonhang` ): chứa danhh sách 500 đơn hàng . Cột A là mã vận đơn , cột D là giá cước cần điền
  -sheet 2 ( `banggiacuoc` ): chứa bảng tra cứu từ ô A2 đến C50
  Công thức tại D2: =VLOOKUP(A2,banggiacuoc!A$2:$C$50,3,FALSE)
## PowerPoimt - slide báo cáo hiệu suất kpi và hướng dẫn shipper
 * Thiết kế slide hướng dẫn shipper nhận đơn giao hàng gồm đúng 3 bước ngắn gọn :
   1. Quét mã QR trên đơn hàng: sử dụng app FastMove driver để nhận diện mã vận đơn trong 1 giây
   2. đối soát giá cước hiển thị
   3. xác nhận xuất kho
# Phần 3. Kiểm thử: - Xử lý lỗi #N/A khi mã vận đơn có ký tự lạ bằng hàm IFERROR.
1. Nguyên nhân xuất hiện lỗi #N/A
Khi nhập Mã vận đơn mới hoặc mã có ký tự lạ/khoảng trắng thừa (ví dụ: FM-99999*), hàm VLOOKUP với tham số FALSE sẽ dò tìm chính xác 100%. Khi không tìm thấy mã này trong BangGiaCuoc, Excel sẽ trả về lỗi #N/A (Not Available).

Lỗi #N/A làm cho toàn bộ các phép tính tổng tiền, tính trung bình ở cuối bảng bị lỗi dây chuyền (#N/A).

2. Xử lý triệt để bằng hàm IFERROR kết hợp TRIM
Công thức lồng IFERROR chuẩn hóa:

Excel
=IFERROR(VLOOKUP(TRIM(A2), BangGiaCuoc!$A$2:$C$50, 3, FALSE), "Chưa có biểu cước")
Quy trình vận hành 3 bước khi phát hiện lỗi:

Bước 1 (Lọc lỗi): Bật AutoFilter (Ctrl + Shift + L), lọc cột Giá cước theo giá trị "Chưa có biểu cước" hoặc chứa #N/A.

Bước 2 (Kiểm tra ký tự lạ / khoảng trắng): Hàm TRIM(A2) đã tự động xóa khoảng trắng thừa. Nếu vẫn báo lỗi, kiểm tra xem Mã vận đơn có bị gõ sai định dạng hoặc chứa ký tự đặc biệt (như dấu *, ?, khoảng trắng không ngắt NBSP) hay không.

Bước 3 (Bổ sung Bảng giá): Nếu mã vận đơn hợp lệ nhưng chưa được đăng ký trong hệ thống, bộ phận kho tiến hành cập nhật mã mới vào BangGiaCuoc. Bảng tính Excel sẽ lập tức tự động cập nhật giá cước mà không cần sửa công thức.
