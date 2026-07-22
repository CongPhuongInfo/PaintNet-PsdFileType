# PsdPlugin - Bản chuyển đổi sang VB.NET (net9.0-windows)

Đây là bản chuyển đổi sang VB.NET của plugin đọc/ghi file PSD/PSB cho
Paint.NET (github.com/PsdPlugin/PsdPlugin), được làm từng phần qua 5 giai
đoạn, build bằng **.NET 9 SDK** — vì bản Paint.NET mới nhất (5.x) đã chuyển
hẳn sang .NET hiện đại (repo gốc hiện target `net7.0-windows`, bản này nâng
lên `net9.0`).

## Cấu trúc thư mục

- `PsdFile/` — thư viện đọc/ghi định dạng PSD thuần túy (header, layer,
  channel, nén ảnh, image resource...), không phụ thuộc Paint.NET
- `PhotoShopFileType/` — phần tích hợp thật vào Paint.NET (đăng ký FileType,
  load/save, decode pixel, dialog cấu hình khi save)
- `Photoshop.vbproj` — file project VB.NET kiểu SDK
- `build.bat` — script build bằng `dotnet build`, tự copy DLL kết quả vào
  thư mục `FileTypes` của Paint.NET luôn

## Cách build

1. Cài [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0) nếu
   chưa có (gõ `dotnet --version` phải ra `9.x`).
2. Mở `Photoshop.vbproj`, kiểm tra biến `PaintNetDir` gần đầu file có đúng
   chỗ cài Paint.NET không — mặc định là `C:\Program Files\paint.net`.
   Project tham chiếu các DLL: `PaintDotNet.Base.dll`, `PaintDotNet.Core.dll`,
   `PaintDotNet.Data.dll`, `PaintDotNet.Primitives.dll`,
   `PaintDotNet.ComponentModel.dll`, `PaintDotNet.Resources.dll`, và
   `System.Drawing.Common.dll` từ thư mục đó. Các tham chiếu này đã đặt
   `Private=False` nên sẽ **không** bị copy kèm vào thư mục output (vì chúng
   đã có sẵn trong thư mục cài Paint.NET rồi).
3. Chạy `build.bat` — file này sẽ:
   - Kiểm tra `dotnet` đã có trên PATH chưa
   - Build project ở cấu hình Release
   - Tự copy `Photoshop.dll` vừa build ra vào
     `C:\Program Files\paint.net\FileTypes\`
   - Nếu Paint.NET đang mở khiến file DLL bị khoá, hoặc không có quyền ghi
     vào `Program Files`, script sẽ báo lỗi rõ ràng để bạn xử lý (đóng
     Paint.NET lại, hoặc chạy cmd với quyền Administrator)
4. Khởi động lại Paint.NET để nó nhận plugin mới.

Nếu muốn build tay không qua `build.bat`: `dotnet build Photoshop.vbproj -c Release`.

### Vì sao `RootNamespace` để trống

Project VB.NET kiểu SDK theo mặc định sẽ tự bọc thêm một "root namespace"
quanh toàn bộ namespace khai báo trong code. Vì mọi file ở đây đã tự khai
báo namespace đầy đủ rồi (`PhotoshopFile`,
`PaintDotNet.Data.PhotoshopFileType`) để khớp đúng với bản C# gốc, nên
`Photoshop.vbproj` đặt `<RootNamespace></RootNamespace>` (rỗng) để tắt việc
tự bọc thêm đó. Nếu bạn build và thấy namespace bị lồng sai (kiểu
`Photoshop.PhotoshopFile` thay vì `PhotoshopFile`), kiểm tra lại chỗ này
trước tiên.

## Những chỗ khác biệt so với bản C# gốc (ghi chú porting)

VB.NET không hỗ trợ `unsafe`/con trỏ, nên mọi chỗ bản gốc dùng
`fixed`/`byte*`/`ColorBgra*` đều được viết lại bằng cách duyệt mảng an toàn:

- Đảo byte (big-endian) trong `PsdBinaryReader`/`Writer`, `Util` → dùng
  `BitConverter` + `Array.Reverse` trên mảng byte.
- Nén/giải nén RLE (`RleReader`/`RleWriter`) → vòng lặp duyệt mảng thường.
- Mã hoá delta kiểu ZIP-with-prediction (`ZipPredict16Image`/`32Image`) →
  dùng `BitConverter.ToUInt16/ToInt32`/`GetBytes` tại đúng vị trí byte, xử
  lý tràn số thủ công bằng `Mod` ở những chỗ bản gốc dựa vào tính chất tràn
  số không kiểm tra (unchecked) của C#.
- Truy cập pixel qua `Surface.GetRowPointer` của Paint.NET
  (`ImageDecoderPdn.vb`, `PsdSave.vb`) → dùng property chỉ số an toàn
  `(x, y)` của `Surface`, giải mã từng dòng ảnh vào một mảng `ColorBgra()`
  tạm (mảng struct nên vẫn sửa trực tiếp từng field được, ví dụ
  `destRow(i).R = ...`, giống hệt `pDest->R = ...` ở bản gốc) rồi mới ghi cả
  dòng ngược lại vào surface một lần.

## Vài lỗi build đã gặp và cách sửa (do SDK Paint.NET mới thay đổi API)

Trong lúc build thử, SDK Paint.NET hiện tại (khác với bản mà PsdPlugin gốc
nhắm tới) có vài thay đổi API khiến phải sửa lại:

- `SaveConfigWidget` giờ là **generic**: `SaveConfigWidget(Of TFileType, TToken)`
  (thay vì lớp không-generic cũ). `PsdSaveConfigWidget` giờ kế thừa
  `SaveConfigWidget(Of PhotoshopFileType, PsdSaveConfigToken)`.
- `InitTokenFromWidget()` giờ bị seal (`NotOverridable`) — không override
  được nữa. Method thật sự cần override là `CreateTokenFromWidget() As TToken`
  (xác nhận qua tài liệu API chính thức tại
  paintdotnet.github.io/apidocs).
- Constructor của `SaveConfigWidget(Of,)` yêu cầu truyền vào một `FileType`:
  `MyBase.New(New PhotoshopFileType())`.

Nếu Paint.NET tiếp tục cập nhật SDK sau này, có thể sẽ còn vài chỗ tương tự
cần chỉnh — cách xác định chính xác nhất là tra trực tiếp
paintdotnet.github.io/apidocs thay vì đoán từ mã nguồn cũ.

## Đã build thành công

Bản này đã build **thành công** với .NET 9 SDK (chỉ còn cảnh báo `CA1416`
về hỗ trợ nền tảng, không phải lỗi). Phần còn lại cần kiểm tra là hành vi
thực tế khi load/save file PSD trong Paint.NET — lỗi runtime (nếu có, ví dụ
sai màu, thiếu layer, sai mask...) sẽ khó đoán trước hơn lỗi biên dịch vì
không có Paint.NET thật để chạy thử trong lúc viết, nên nếu gặp báo lại để
soát lại logic tương ứng.
