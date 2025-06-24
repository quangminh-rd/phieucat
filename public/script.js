var DocTienBangChu = function () {
    this.ChuSo = new Array(" không ", " một ", " hai ", " ba ", " bốn ", " năm ", " sáu ", " bảy ", " tám ", " chín ");
    this.Tien = new Array("", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ");
};

DocTienBangChu.prototype.docSo3ChuSo = function (baso) {
    var tram;
    var chuc;
    var donvi;
    var KetQua = "";
    tram = parseInt(baso / 100);
    chuc = parseInt((baso % 100) / 10);
    donvi = baso % 10;
    if (tram == 0 && chuc == 0 && donvi == 0) return "";
    if (tram != 0) {
        KetQua += this.ChuSo[tram] + " trăm ";
        if ((chuc == 0) && (donvi != 0)) KetQua += " linh ";
    }
    if ((chuc != 0) && (chuc != 1)) {
        KetQua += this.ChuSo[chuc] + " mươi";
        if ((chuc == 0) && (donvi != 0)) KetQua = KetQua + " linh ";
    }
    if (chuc == 1) KetQua += " mười ";
    switch (donvi) {
        case 1:
            if ((chuc != 0) && (chuc != 1)) {
                KetQua += " mốt ";
            }
            else {
                KetQua += this.ChuSo[donvi];
            }
            break;
        case 5:
            if (chuc == 0) {
                KetQua += this.ChuSo[donvi];
            }
            else {
                KetQua += " lăm ";
            }
            break;
        default:
            if (donvi != 0) {
                KetQua += this.ChuSo[donvi];
            }
            break;
    }
    return KetQua;
}

DocTienBangChu.prototype.doc = function (SoTien) {
    var lan = 0;
    var i = 0;
    var so = 0;
    var KetQua = "";
    var tmp = "";
    var soAm = false;
    var ViTri = new Array();
    if (SoTien < 0) soAm = true;//return "Số tiền âm !";
    if (SoTien == 0) return "Không đồng";//"Không đồng !";
    if (SoTien > 0) {
        so = SoTien;
    }
    else {
        so = -SoTien;
    }
    if (SoTien > 8999999999999999) {
        //SoTien = 0;
        return "";//"Số quá lớn!";
    }
    ViTri[5] = Math.floor(so / 1000000000000000);
    if (isNaN(ViTri[5]))
        ViTri[5] = "0";
    so = so - parseFloat(ViTri[5].toString()) * 1000000000000000;
    ViTri[4] = Math.floor(so / 1000000000000);
    if (isNaN(ViTri[4]))
        ViTri[4] = "0";
    so = so - parseFloat(ViTri[4].toString()) * 1000000000000;
    ViTri[3] = Math.floor(so / 1000000000);
    if (isNaN(ViTri[3]))
        ViTri[3] = "0";
    so = so - parseFloat(ViTri[3].toString()) * 1000000000;
    ViTri[2] = parseInt(so / 1000000);
    if (isNaN(ViTri[2]))
        ViTri[2] = "0";
    ViTri[1] = parseInt((so % 1000000) / 1000);
    if (isNaN(ViTri[1]))
        ViTri[1] = "0";
    ViTri[0] = parseInt(so % 1000);
    if (isNaN(ViTri[0]))
        ViTri[0] = "0";
    if (ViTri[5] > 0) {
        lan = 5;
    }
    else if (ViTri[4] > 0) {
        lan = 4;
    }
    else if (ViTri[3] > 0) {
        lan = 3;
    }
    else if (ViTri[2] > 0) {
        lan = 2;
    }
    else if (ViTri[1] > 0) {
        lan = 1;
    }
    else {
        lan = 0;
    }
    for (i = lan; i >= 0; i--) {
        tmp = this.docSo3ChuSo(ViTri[i]);
        KetQua += tmp;
        if (ViTri[i] > 0) KetQua += this.Tien[i];
        if ((i > 0) && (tmp.length > 0)) KetQua += '';//',';//&& (!string.IsNullOrEmpty(tmp))
    }
    if (KetQua.substring(KetQua.length - 1) == ',') {
        KetQua = KetQua.substring(0, KetQua.length - 1);
    }
    KetQua = KetQua.substring(1, 2).toUpperCase() + KetQua.substring(2);
    if (soAm) {
        return "Âm " + KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
    else {
        return KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
}

function formatNumber(numberString) {
    if (!numberString) return '';
    // Loại bỏ tất cả dấu chấm
    const num = numberString.replace(/\./g, '');
    const formatted = parseFloat(num).toString();
    return formatted.replace('.', ',');
}

const SPREADSHEET_ID = '14R9efcJ2hGE3mCgmJqi6TNbqkm4GFe91LEAuCyCa4O0';
const RANGE = 'don_hang!A:BO'; // Mở rộng phạm vi đến cột BO
const RANGE_CHITIET = 'don_hang_chi_tiet!F:FZ'; // Dải dữ liệu từ sheet 'don_hang_chi_tiet'


const SPREADSHEET_ID_TENTRUONG = '1gY6a0TWrXeQhLrMT0TLuGpISlU8VUoA6pl_Rq9PQ0xU';
const RANGE_LOOKUP = 'ten_truong_chon!A:AL'; // Phạm vi chứa dữ liệu mapping

const SPREADSHEET_ID_VATTU = '1FU4f9p6LPUCf_7bMVmS6kWJ7n7kA5_hDSqkUO1sGqdQ';
const RANGE_VATTU = 'dsvt!F:G'; // Cột F: mã vật tư, cột G: tên vật tư

let materialTable = [];


const API_KEY = 'AIzaSyA9g2qFUolpsu3_HVHOebdZb0NXnQgXlFM';

// Lấy giá trị từ URI sau dấu "?" cho các tham số cụ thể
function getDataFromURI() {
    const url = window.location.href;

    // Sử dụng RegEx để trích xuất giá trị của ma_don_hang và QRCODE
    const maDonHangURIMatch = url.match(/ma_don_hang=([^?&]*)/);
    const maKhachHangURIMatch = url.match(/ma_khach_hang=([^?&]*)/);
    const qrCodeMatch = url.match(/QRCODE=(.*)$/);  // Sử dụng regex để lấy tất cả sau QRCODE=

    // Gán các giá trị vào các biến
    const maDonHangURI = maDonHangURIMatch ? decodeURIComponent(maDonHangURIMatch[1]) : null;
    const maKhachHangURI = maKhachHangURIMatch ? decodeURIComponent(maKhachHangURIMatch[1]) : null;
    const qrCode = qrCodeMatch ? decodeURIComponent(qrCodeMatch[1]) : null;

    // Trả về một đối tượng chứa các giá trị
    return {
        maDonHangURI,
        maKhachHangURI,
        qrCode
    };
}

// Hàm để tải Google API Client
function loadGapiAndInitialize() {
    const script = document.createElement('script');
    script.src = "https://apis.google.com/js/api.js"; // Đường dẫn đến Google API Client
    script.onload = initialize; // Gọi hàm `initialize` sau khi thư viện được tải xong
    script.onerror = () => console.error('Failed to load Google API Client.');
    document.body.appendChild(script); // Gắn thẻ script vào tài liệu
}

// Hàm khởi tạo sau khi Google API Client được tải
function initialize() {
    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });

            const uriData = getDataFromURI();
            if (!uriData.maDonHangURI || !uriData.qrCode) {
                updateContent('No valid data found in URI.');
                return;
            }

            await loadLookupTable();
            findRowInSheet(uriData.maDonHangURI);
            findDetailsInSheet(uriData.maDonHangURI);
            loadMaterialTable();

            // Cập nhật nội dung hoặc xử lý thêm thông tin QR Code
            updateQRCodeContent(uriData.qrCode);

        } catch (error) {
            updateContent('Initialization error: ' + error.message);
            console.error('Initialization Error:', error);
        }
    });
}

// Gọi hàm tải Google API Client khi DOM đã sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    loadGapiAndInitialize();
});

function updateQRCodeContent(qrCode) {
    // Gắn QR Code vào trong nội dung trang (VD: hiển thị ảnh QR code)
    const qrCodeElement = document.getElementById('qr-code');
    if (qrCodeElement) {
        qrCodeElement.src = qrCode;
        // Đặt kích thước của QR Code
        qrCodeElement.style.width = '300px';  // Chiều rộng 150px
        qrCodeElement.style.height = 'auto';  // Chiều cao tự động
    }
}

function updateContent(message) {
    const contentElement = document.getElementById('content'); // Thay 'content' bằng ID của phần tử HTML cần hiển thị
    if (contentElement) {
        contentElement.textContent = message;
    } else {
        console.warn('Element with ID "content" not found.');
    }
}

let lookupTable = [];

function loadLookupTable() {
    gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID_TENTRUONG,
        range: RANGE_LOOKUP,
    }).then(response => {
        const values = response.result.values;
        if (!values || values.length < 2) return;

        const headers = values[0];
        lookupTable = values.slice(1).map(row => {
            const obj = {};
            headers.forEach((header, idx) => {
                obj[header.trim()] = row[idx] || '';
            });
            return obj;
        });
    }).catch(error => {
        console.error('Failed to load lookup table:', error);
    });
}

function getFieldValue(maSanpham, fieldName) {
    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    return row ? row[fieldName] || '' : '';
}

function getFieldValue_truong1(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_1'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong2(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_2'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong3(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_3'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong4(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_4'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong5(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_5'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong6(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_6'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong7(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_7'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong8(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_8'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong9(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_9'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValue_truong10(maSanpham) {
    if (!maSanpham) return '';

    const row = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!row) return '';

    const dynamicFieldName = row['truong_10'];
    if (!dynamicFieldName) return '';

    return row[dynamicFieldName] || '';
}

function getFieldValueFromTruong(maSanpham, truongKey) {
    if (!maSanpham) return '';

    const lookupRow = lookupTable.find(item => item['ma_san_pham'] === maSanpham);
    if (!lookupRow || !lookupRow[truongKey]) return '';

    const fieldName = lookupRow[truongKey];
    const detailRow = detailTable.find(item => item['ma_san_pham_cau_tao'] === maSanpham);
    if (!detailRow || !detailRow[fieldName]) return '';

    return detailRow[fieldName];
}

function loadMaterialTable() {
    gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID_VATTU,
        range: RANGE_VATTU,
    }).then(response => {
        const values = response.result.values;
        if (!values || values.length < 2) return;

        materialTable = values.slice(1).map(row => ({
            maVatTu: row[0]?.trim() || '',
            tenVatTu: row[1]?.trim() || ''
        }));
    }).catch(error => {
        console.error('[ERROR] loadMaterialTable:', error);
    });
}

function getTenVatTu(maVatTu) {
    if (!maVatTu) return '';
    const item = materialTable.find(vt => vt.maVatTu === maVatTu);
    return item ? item.tenVatTu : '';
}

// Tìm chỉ số dòng chứa dữ liệu khớp trong cột B và lấy các giá trị từ các cột khác
let orderDetails = null; // Thông tin đơn hàng chính
let orderItems = [];

async function findRowInSheet(maDonhangURI) {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            updateContent('No data found.');
            return;
        }

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            const bColumnValue = row[1]; // Cột B
            const uriData = getDataFromURI();
            if (bColumnValue === maDonhangURI) {
                // Lưu dữ liệu vào biến toàn cục
                orderDetails = {
                    phuongThucban: row[0] || '', // Cột A
                    maDonhang: row[1] || '', // Cột B
                    maHopdong: row[66] || '', // Cột BO
                    donviPhutrach: row[5] || '', // Cột F
                    tenNguoilienhe: row[13] || '', // Cột N
                    tenKhachhangcuoi: row[23] || '', // Cột X
                    tenTochuc: row[15] || '', // Cột P
                    diachi: row[8] || '', // Cột I
                    diachiChitiet: row[17] || '', // Cột R
                    diachiKhachhangcuoi: row[24] || '', // Cột Y
                    tenNhanvien: row[4] || '', // Cột E
                    sdtNhanvien: row[6] || '', // Cột G
                    sdtKhachhang: row[18] || '', // Cột S
                    sdtKhachhangcuoi: row[22] || '', // Cột W
                    emailKhachhang: row[19] || '', // Cột T
                    hanGiaohang: row[70] || '', // Cột BS
                    tongSobo: row[25] || '', // Cột Z
                    congnpp: row[43] || '', // Cột AR
                    mucChietkhaunpp: row[44] || '', // Cột AS
                    giatriChietkhaunpp: row[45] || '', // Cột AT
                    phiVanchuyenlapdatnpp: row[46] || '', // Cột AU
                    mucthueGTGTnpp: row[47] || '', // Cột AV
                    thueGTGTnpp: row[48] || '', // Cột AW
                    tamUngnpp: row[49] || '', // Cột AX
                    sotienConthieunpp: row[58] || '',
                    maKhachHang: uriData.maKhachHangURI,
                };

                // Xử lý dữ liệu tìm được
                processFoundData(orderDetails);
                return; // Dừng khi tìm thấy
            }
        }

        updateContent(`No matching data found for "${maDonhangURI}".`);
    } catch (error) {
        updateContent('Error fetching data: ' + error.message);
        console.error('Fetch Error:', error);
    }
}

function processFoundData(orderDetails) {
    // Định dạng và chuyển đổi số tiền
    const formattedSotien = formatNumber(orderDetails.sotienConthieunpp || '0');
    const doctien = new DocTienBangChu();
    const sotienBangchu = doctien.doc(formattedSotien);

    // Cập nhật giá trị sotienBangchu vào orderDetails
    orderDetails.sotienBangchu = sotienBangchu;

    // Cập nhật DOM
    Object.keys(orderDetails).forEach((key) => {
        if (orderDetails[key]) {
            updateElement(key, orderDetails[key]);
        }
    });
    if (sotienBangchu) updateElement('sotienBangchu', sotienBangchu);

    // Gọi các hàm hiển thị nội dung
    displayHTML(orderDetails);

    function toggleRowVisibility(rowId, value) {
        const row = document.getElementById(rowId);
        if (row) {
            const stringValue = typeof value === 'string' ? value : String(value || '');
            const numericValue = parseFloat(stringValue.replace(/\./g, '').replace(',', '.') || '0');
            row.style.display = numericValue > 0 ? '' : 'none';
        }
    }

    // Ẩn/hiện các dòng theo điều kiện
    toggleRowVisibility('rowChietKhau', orderDetails.giatriChietkhaunpp);
    toggleRowVisibility('rowPhiVanChuyen', orderDetails.phiVanchuyenlapdatnpp);
    toggleRowVisibility('rowThueGTGT', orderDetails.thueGTGTnpp);
    toggleRowVisibility('rowTamUng', orderDetails.tamUngnpp);

    // Hiển thị hoặc ẩn nội dung thanh toán
    const paymentContent = document.getElementById('payment-content');
    if (paymentContent) {
        paymentContent.style.display =
            orderDetails.donviPhutrach === "BP. BH1" || orderDetails.donviPhutrach === "BP. BH2"
                ? 'block'
                : 'none';
    }
}

function displayHTML() {
    // Trích xuất các giá trị cần thiết từ data
    const maDonhang = orderDetails.maDonhang || '';
    const phuongThucban = orderDetails.phuongThucban || '';
    const donviPhutrach = orderDetails.donviPhutrach || '';
    const tenNguoilienhe = orderDetails.tenNguoilienhe || '';
    const tenKhachhangcuoi = orderDetails.tenKhachhangcuoi || '';
    const tenTochuc = orderDetails.tenTochuc || '';
    const diachi = orderDetails.diachi || '';
    const diachiChitiet = orderDetails.diachiChitiet || '';
    const diachiKhachhangcuoi = orderDetails.diachiKhachhangcuoi || '';
    const tenNhanvien = orderDetails.tenNhanvien || '';
    const sdtNhanvien = orderDetails.sdtNhanvien || '';
    const sdtKhachhang = orderDetails.sdtKhachhang || '';
    const sdtKhachhangcuoi = orderDetails.sdtKhachhangcuoi || '';
    const emailKhachhang = orderDetails.emailKhachhang || '';
    const hanGiaohang = orderDetails.hanGiaohang || '';
    const today = new Date();
    const ngayPhatHanh = today.toLocaleDateString('vi-VN');
    if (ngayPhatHanh) updateElement('ngayPhatHanh', ngayPhatHanh);
    // Cập nhật giá trị ngayPhatHanh vào orderDetails
    orderDetails.ngayPhatHanh = ngayPhatHanh;


    let htmlContent = "";

    if (donviPhutrach === "BP. BH1" && phuongThucban !== "Bán chéo") {
        htmlContent = `
                            <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Ngày:</i></td>
                                        <td class="infocol-2">${ngayPhatHanh}</td>
                                        <td class="infocol-3"><i>Mã báo giá:</i></td>
                                        <td class="infocol-4">${maDonhang}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Ngày sản xuất:</i></td>
                                        <td class="infocol-2"></td>
                                        <td class="infocol-3"><i>Khách hàng:</i></td>
                                        <td class="infocol-4">${tenNguoilienhe} - ${diachiChitiet} - KD: ${tenNhanvien}</td>
                                    </tr>
                                </tbody>
                            `;
    } else if (donviPhutrach === "BP. BH1" && phuongThucban === "Bán chéo") {
        htmlContent = `
                                <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Ngày:</i></td>
                                        <td class="infocol-2">${ngayPhatHanh}</td>
                                        <td class="infocol-3"><i>Mã báo giá:</i></td>
                                        <td class="infocol-4">${maDonhang}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Ngày sản xuất:</i></td>
                                        <td class="infocol-2"></td>
                                        <td class="infocol-3"><i>Khách hàng:</i></td>
                                        <td class="infocol-4">${tenKhachhangcuoi} - ${diachiKhachhangcuoi} - KD: ${tenNhanvien}</td>
                                    </tr>
                                </tbody>
                                `;
    } else if (donviPhutrach !== "BP. BH1") {
        htmlContent = `
                                   <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Ngày:</i></td>
                                        <td class="infocol-2">${ngayPhatHanh}</td>
                                        <td class="infocol-3"><i>Mã báo giá:</i></td>
                                        <td class="infocol-4">${maDonhang}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Ngày sản xuất:</i></td>
                                        <td class="infocol-2"></td>
                                        <td class="infocol-3"><i>Khách hàng:</i></td>
                                        <td class="infocol-4">${donviPhutrach}</td>
                                    </tr>
                                </tbody>
                                `;
    }

    document.getElementById("content").innerHTML = htmlContent;
}

async function findDetailsInSheet(maDonhangURI) {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: RANGE_CHITIET,
        });

        const allRows = response.result.values;
        if (!allRows || allRows.length === 0) {
            updateContent('No detail data found.');
            return;
        }

        const headers = allRows[0]; // Tiêu đề
        const dataRows = allRows.slice(1); // Dữ liệu bỏ dòng tiêu đề

        const filteredRows = dataRows.filter(row => row[0] === maDonhangURI); // Cột F
        orderItems = filteredRows.map(extractDetailDataFromRow);

        // Tạo detailTable đúng cách
        detailTable = filteredRows.map(row => {
            const obj = {};
            headers.forEach((header, idx) => {
                obj[header.trim()] = row[idx] || '';
            });
            return obj;
        });

        console.log('[DEBUG] Số dòng filteredRows:', filteredRows.length);
        console.log('[DEBUG] detailTable sample:', detailTable.slice(0, 2));
        console.log('[DEBUG] Các mã có trong detailTable:', detailTable.map(row => row['ma_san_pham_id']).slice(0, 10));

        if (filteredRows.length > 0) {
            displayDetailData(filteredRows);
        } else {
            updateContent('No matching data found.');
        }
    } catch (error) {
        console.error('Error fetching detail data:', error);
        updateContent('Error fetching detail data.');
    }
}

function displayDetailData(filteredRows) {
    const tableBody = document.getElementById('itemTableBody');
    tableBody.innerHTML = ''; // Xóa dữ liệu cũ

    // Lọc và chuẩn hóa dữ liệu (bỏ xử lý nhóm)
    const items = filteredRows.map(extractDetailDataFromRow);

    items.forEach(item => {
        // Xử lý hiển thị kích thước theo đơn vị tính
        let chieuRongHTML = '', chieuCaoHTML = '', dienTichTencot = ''; dienTichHTML = '';

        if (item.dvt === 'm2' || item.dvt === 'bộ') {
            chieuRongHTML = `<td class="borderedcol-9">${item.chieuRong || ''}</td>`;
            chieuCaoHTML = `<td class="borderedcol-10">${item.chieuCao || ''}</td>`;
            dienTichTencot = `<th class="borderedcol-11">Diện tích<br>(m2)</th>`;
            dienTichHTML = `<td class="borderedcol-11">${item.dienTich || ''}</td>`;
        } else if (item.dvt === 'm') {
            chieuRongHTML = `<td class="borderedcol-9"></td>`;
            chieuCaoHTML = `<td class="borderedcol-10">${item.chieuCao || ''}</td>`;
            dienTichTencot = `<th class="borderedcol-11"></th>`;
            dienTichHTML = `<td class="borderedcol-11"></td>`;
        } else {
            chieuRongHTML = `<td class="borderedcol-9"></td>`;
            chieuCaoHTML = `<td class="borderedcol-10"></td>`;
            dienTichTencot = `<th class="borderedcol-11"></th>`;
            dienTichHTML = `<td class="borderedcol-11"></td>`;
        }

        const diengiaiGhiChu = !item.maDonhangCT.includes("1C.029.01") && item.ghiChu
            ? `${item.diengiai || ''} - ${item.ghiChu}`
            : item.diengiai || '';

        let mvt_HTML = ``;

        for (let i = 1; i <= 30; i++) {
            const ma_mvt = item[`mvt${i}`];
            const ten_mvt = getTenVatTu(ma_mvt);
            const kt_mvt = item[`ktMvt${i}`];
            const sl_mvt = item[`slMvt${i}`];

            if (!ma_mvt) continue;

            mvt_HTML += `
                    <tr class="bordered-table">
                        <td class="borderedcol-1">${i}</td>
                        <td class="borderedcol-2-body">${ma_mvt}</td>
                        <td class="borderedcol-3-body" colspan="3">${ten_mvt}</td>
                        <td class="borderedcol-6">${kt_mvt || ''}</td>
                        <td class="borderedcol-7">${sl_mvt || ''}</td>
                        <td class="borderedcol-8"></td>
                        <td class="borderedcol-9"></td>
                        <td class="borderedcol-10"></td>
                        <td class="borderedcol-11"></td>
                        <td class="borderedcol-12"></td>
                    </tr>
                    `;
        }

        let m2vt_HTML = ``;

        for (let i = 1; i <= 30; i++) {
            const ma_m2vt = item[`m2vt${i}`];
            const ten_m2vt = getTenVatTu(ma_m2vt);
            const kt1_m2vt = item[`kt1M2vt${i}`];
            const kt2_m2vt = item[`kt2M2vt${i}`];
            const sl_m2vt = item[`slM2vt${i}`];

            if (!ma_m2vt) continue;

            m2vt_HTML += `
                    <tr class="bordered-table">
                        <td class="borderedcol-1">${i}</td>
                        <td class="borderedcol-2-body">${ma_m2vt}</td>
                        <td class="borderedcol-3-body" colspan="2">${ten_m2vt}</td>
                        <td class="borderedcol-5">${kt1_m2vt || ''}</td>
                        <td class="borderedcol-6">${kt2_m2vt || ''}</td>
                        <td class="borderedcol-7">${sl_m2vt || ''}</td>
                        <td class="borderedcol-8"></td>
                        <td class="borderedcol-9"></td>
                        <td class="borderedcol-10"></td>
                        <td class="borderedcol-11"></td>
                        <td class="borderedcol-12"></td>
                    </tr>
                    `;
        }

        let cvt_HTML = ``;

        for (let i = 1; i <= 30; i++) {
            const ma_cvt = item[`cvt${i}`];
            const ten_cvt = getTenVatTu(ma_cvt);
            const sl_cvt = item[`slCvt${i}`];

            if (!ma_cvt) continue;

            cvt_HTML += `
                    <tr class="bordered-table">
                        <td class="borderedcol-1">${i}</td>
                        <td class="borderedcol-2-body">${ma_cvt}</td>
                        <td class="borderedcol-3-body" colspan="4">${ten_cvt}</td>
                        <td class="borderedcol-7">${sl_cvt || ''}</td>
                        <td class="borderedcol-8"></td>
                        <td class="borderedcol-9"></td>
                        <td class="borderedcol-10"></td>
                        <td class="borderedcol-11"></td>
                        <td class="borderedcol-12"></td>
                    </tr>
                    `;
        }

        tableBody.innerHTML += `
        <tr class="bordered-table" style="background-color: #A4CBE9;">
            <th class="borderedcol-1">STT</th>
            <th class="borderedcol-2">Mã sản phẩm</th>
            <th class="borderedcol-3">${getFieldValue(item.maSanphamid, 'tong_so_canh')}</th>
            <th class="borderedcol-4">${getFieldValue(item.maSanphamid, 'khung_nhom')}</th>
            <th class="borderedcol-5" colspan = 2 >${getFieldValue(item.maSanphamid, 'mau_luoi')}</th>
            <th class="borderedcol-7" colspan = 2 >${getFieldValue(item.maSanphamid, 'mau_rem')}</th>
            <th class="borderedcol-9">${getFieldValue(item.maSanphamid, 'chieu_rong')}</th>
            <th class="borderedcol-10">${getFieldValue(item.maSanphamid, 'chieu_cao')}</th>
            ${dienTichTencot}
            <th class="borderedcol-12">Số lượng</th>
        </tr>
        <tr class="bordered-table">
            <th class="borderedcol-1" rowspan = 3 style="color: red; background-color: yellow;">${item.sttTrongdon || ''}</th>
            <td class="borderedcol-2">${item.maSanphamid || ''}</td>
            <td class="borderedcol-3">${item.tongSoCanh || ''}</td>
            <td class="borderedcol-4">${item.khungNhom || ''}</td>
            <td class="borderedcol-5" colspan = 2 >${item.mauLuoi || ''}</td>
            <td class="borderedcol-7" colspan = 2 >${item.mauRem || ''}</td>
            ${chieuRongHTML}
            ${chieuCaoHTML}
            ${dienTichHTML}
            <td class="borderedcol-12">${item.soLuong || ''}</td>
        </tr>
        <tr class="bordered-table" style="background-color: #A4CBE9;">
            <th class="borderedcol-2">${getFieldValue_truong1(item.maSanphamid)}</th>
            <th class="borderedcol-3">${getFieldValue_truong2(item.maSanphamid)}</th>
            <th class="borderedcol-4">${getFieldValue_truong3(item.maSanphamid)}</th>
            <th class="borderedcol-5" colspan = 2 >${getFieldValue_truong4(item.maSanphamid)}</th>
            <th class="borderedcol-7" colspan = 2 >${getFieldValue_truong5(item.maSanphamid)}</th>
            <th class="borderedcol-9">${getFieldValue_truong6(item.maSanphamid)}</th>
            <th class="borderedcol-10">${getFieldValue_truong7(item.maSanphamid)}</th>
            <th class="borderedcol-11">${getFieldValue_truong8(item.maSanphamid)}</th>
            <th class="borderedcol-12">${getFieldValue_truong9(item.maSanphamid)}</th>
        </tr>
        <tr class="bordered-table">
            <td class="borderedcol-2">${getFieldValueFromTruong(item.maSanphamid, 'truong_1')}</td>
            <td class="borderedcol-3">${getFieldValueFromTruong(item.maSanphamid, 'truong_2')}</td>
            <td class="borderedcol-4">${getFieldValueFromTruong(item.maSanphamid, 'truong_3')}</td>
            <td class="borderedcol-5" colspan = 2 >${getFieldValueFromTruong(item.maSanphamid, 'truong_4')}</td>
            <td class="borderedcol-7" colspan = 2 >${getFieldValueFromTruong(item.maSanphamid, 'truong_5')}</td>
            <td class="borderedcol-9">${getFieldValueFromTruong(item.maSanphamid, 'truong_6')}</td>
            <td class="borderedcol-10">${getFieldValueFromTruong(item.maSanphamid, 'truong_7')}</td>
            <td class="borderedcol-11">${getFieldValueFromTruong(item.maSanphamid, 'truong_8')}</td>
            <td class="borderedcol-12">${getFieldValueFromTruong(item.maSanphamid, 'truong_9')}</td>
        </tr>
        <tr class="bordered-table" style="background-color: #F4C392;">
            <th class="borderedcol-1">A</th>
            <th class="borderedcol-2">Mã vật tư</th>
            <th class="borderedcol-3" colspan = 3 >Tên vật tư</th>
            <th class="borderedcol-6">Chiều dài</th>
            <th class="borderedcol-7">Số lượng</th>
            <th class="borderedcol-8">Cắt nhôm</th>
            <th class="borderedcol-9">Hoàn thiện</th>
            <th class="borderedcol-10">Kiểm thử</th>
            <th class="borderedcol-11">Xác nhận</th>
            <th class="borderedcol-12">Ghi chú</th>
        </tr>
        ${mvt_HTML}
        <tr class="bordered-table" style="background-color: #F4C392;">
            <th class="borderedcol-1">B</th>
            <th class="borderedcol-2">Mã vật tư</th>
            <th class="borderedcol-3" colspan = 2 >Tên vật tư</th>
            <th class="borderedcol-5">Chiều rộng<br>(Số nếp)</th>
            <th class="borderedcol-6">Chiều dài</th>
            <th class="borderedcol-7">Số lượng</th>
            <th class="borderedcol-8">Cắt nhôm</th>
            <th class="borderedcol-9">Hoàn thiện</th>
            <th class="borderedcol-10">Kiểm thử</th>
            <th class="borderedcol-11">Xác nhận</th>
            <th class="borderedcol-12">Ghi chú</th>
        </tr>
        ${m2vt_HTML}
        <tr class="bordered-table" style="background-color: #F4C392;">
            <th class="borderedcol-1">C</th>
            <th class="borderedcol-2">Mã vật tư</th>
            <th class="borderedcol-3" colspan = 4 >Tên vật tư</th>
            <th class="borderedcol-7">Số lượng</th>
            <th class="borderedcol-8">Cắt nhôm</th>
            <th class="borderedcol-9">Hoàn thiện</th>
            <th class="borderedcol-10">Kiểm thử</th>
            <th class="borderedcol-11">Xác nhận</th>
            <th class="borderedcol-12">Ghi chú</th>
        </tr>
        ${cvt_HTML}
        <tr class="bordered-table">
            <th class="borderedcol-1" colspan = 12 ></th>
        </tr>
        `;
    });
}

// Trích xuất dữ liệu từ hàng
function extractDetailDataFromRow(row) {
    return {
        maDonhangCT: row[0],
        group: row[1],
        sttTrongdon: row[2],
        maSanphamid: row[10],
        ghiChu: row[19],
        chieuRong: row[12],
        chieuCao: row[13],
        dienTich: row[14],
        soLuong: row[16],
        dvt: row[17],
        kieuCua: row[30],
        khungNhom: row[31],
        tongSoCanh: row[32],
        soCanhDiDong: row[33],
        soCanhLuoi: row[34],
        soCanhRem: row[35],
        mauLuoi: row[36],
        mauRem: row[37],
        soCanhBenTrai: row[38],
        soCanhBenPhai: row[39],
        kieuTruot: row[40],
        soDoNgang: row[41],
        soDoDoc: row[42],
        kieuCanh: row[43],
        soKhungDung: row[44],
        rayLuaKhacChuan: row[45],
        mauNoiLa: row[46],
        huongLoiLaNhom: row[47],
        phanTramTangRem: row[48],
        mvt1: row[49],
        ktMvt1: row[50],
        slMvt1: row[51],
        mvt2: row[52],
        ktMvt2: row[53],
        slMvt2: row[54],
        mvt3: row[55],
        ktMvt3: row[56],
        slMvt3: row[57],
        mvt4: row[58],
        ktMvt4: row[59],
        slMvt4: row[60],
        mvt5: row[61],
        ktMvt5: row[62],
        slMvt5: row[63],
        mvt6: row[64],
        ktMvt6: row[65],
        slMvt6: row[66],
        mvt7: row[67],
        ktMvt7: row[68],
        slMvt7: row[69],
        mvt8: row[70],
        ktMvt8: row[71],
        slMvt8: row[72],
        mvt9: row[73],
        ktMvt9: row[74],
        slMvt9: row[75],
        mvt10: row[76],
        ktMvt10: row[77],
        slMvt10: row[78],
        mvt11: row[79],
        ktMvt11: row[80],
        slMvt11: row[81],
        mvt12: row[82],
        ktMvt12: row[83],
        slMvt12: row[84],
        mvt13: row[85],
        ktMvt13: row[86],
        slMvt13: row[87],
        mvt14: row[88],
        ktMvt14: row[89],
        slMvt14: row[90],
        mvt15: row[91],
        ktMvt15: row[92],
        slMvt15: row[93],
        mvt16: row[94],
        ktMvt16: row[95],
        slMvt16: row[96],
        mvt17: row[97],
        ktMvt17: row[98],
        slMvt17: row[99],
        mvt18: row[100],
        ktMvt18: row[101],
        slMvt18: row[102],
        mvt19: row[103],
        ktMvt19: row[104],
        slMvt19: row[105],
        mvt20: row[106],
        ktMvt20: row[107],
        slMvt20: row[108],
        m2vt1: row[109],
        kt1M2vt1: row[110],
        kt2M2vt1: row[111],
        slM2vt1: row[112],
        m2vt2: row[113],
        kt1M2vt2: row[114],
        kt2M2vt2: row[115],
        slM2vt2: row[116],
        cvt1: row[117],
        slCvt1: row[118],
        cvt2: row[119],
        slCvt2: row[120],
        cvt3: row[121],
        slCvt3: row[122],
        cvt4: row[123],
        slCvt4: row[124],
        cvt5: row[125],
        slCvt5: row[126],
        cvt6: row[127],
        slCvt6: row[128],
        cvt7: row[129],
        slCvt7: row[130],
        cvt8: row[131],
        slCvt8: row[132],
        cvt9: row[133],
        slCvt9: row[134],
        cvt10: row[135],
        slCvt10: row[136],
        cvt11: row[137],
        slCvt11: row[138],
        cvt12: row[139],
        slCvt12: row[140],
        cvt13: row[141],
        slCvt13: row[142],
        cvt14: row[143],
        slCvt14: row[144],
        cvt15: row[145],
        slCvt15: row[146],
        cvt16: row[147],
        slCvt16: row[148],
        cvt17: row[149],
        slCvt17: row[150],
        cvt18: row[151],
        slCvt18: row[152],
        cvt19: row[153],
        slCvt19: row[154],
        cvt20: row[155],
        slCvt20: row[156],
        cvt21: row[157],
        slCvt21: row[158],
        cvt22: row[159],
        slCvt22: row[160],
        cvt23: row[161],
        slCvt23: row[162],
        cvt24: row[163],
        slCvt24: row[164],
        cvt25: row[165],
        slCvt25: row[166],
        cvt26: row[167],
        slCvt26: row[168],
        cvt27: row[169],
        slCvt27: row[170],
        cvt28: row[171],
        slCvt28: row[172],
        cvt29: row[173],
        slCvt29: row[174],
        cvt30: row[175],
        slCvt30: row[176],

    };
}

// Hàm cập nhật nội dung DOM
function updateElement(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerText = value;
    }
}