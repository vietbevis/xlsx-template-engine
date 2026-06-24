import { mkdir } from 'node:fs/promises';
import { defineWorkbook } from '../core/workbook';
import { renderWorkbook } from '../renderer';

const danhSachDoiTuong = [
  { title: 'VN', id: 'vn' },
  { title: 'Lào', id: 'lao' },
  { title: 'Cuba', id: 'cuba' },
  { title: 'CPC', id: 'cpc' },
  { title: 'Đóng HP', id: 'dong_hp' },
];

const columns = [
  { title: 'STT', id: 'stt', bodyStyle: { alignment: { horizontal: 'center' } } },
  {
    title: 'Họ tên giảng viên',
    id: 'giangVien',
    bodyStyle: { alignment: { horizontal: 'left', wrapText: true } },
    width: 25,
  },
  {
    title: 'Thu nhập (lương thực nhận, không tính phụ cấp học hàm, học vị)',
    id: 'thuNhap',
    bodyStyle: 'money',
    width: 16,
  },
  { title: 'Định mức giờ giảng', id: 'dinhMucGioGiang', bodyStyle: 'number' },
  { title: 'Được giảm', id: 'duocGiam', bodyStyle: 'number' },
  {
    title: 'Số tiết chưa hoàn thành NCKH',
    id: 'soTietChuaHoanThanhNCKH',
    bodyStyle: 'number',
  },
  { title: 'Định mức phải giảng', id: 'dinhMucPhaiGiang', bodyStyle: 'number' },
  {
    title: 'Thực tế giảng dạy',
    children: [
      {
        title: 'Học kỳ I',
        children: danhSachDoiTuong.map((doiTuong) => ({
          title: doiTuong.title,
          id: `hocKyI_${doiTuong.id}`,
          bodyStyle: 'number',
        })),
      },
      {
        title: 'Học kỳ II',
        children: danhSachDoiTuong.map((doiTuong) => ({
          title: doiTuong.title,
          id: `hocKyII_${doiTuong.id}`,
          bodyStyle: 'number',
        })),
      },
      {
        title: 'Cả năm',
        children: Array.from(
          [...danhSachDoiTuong, { title: 'Tổng', id: 'tong' }].map((doiTuong) => ({
            title: doiTuong.title,
            id: `caNam_${doiTuong.id}`,
            bodyStyle: 'number',
          })),
        ),
      },
    ],
  },
  { title: 'Số tiết vượt định mức', id: 'soTietVuotDinhMuc', bodyStyle: 'number' },
  { title: 'Mức TT chuẩn', id: 'mucTTChuan', bodyStyle: 'money' },
  {
    title: 'Thành tiền',
    childrenRowOffset: 2,
    children: Array.from(
      [...danhSachDoiTuong, { title: 'Tổng', id: 'tong' }].map((doiTuong) => ({
        ...doiTuong,
        bodyStyle: 'money',
        width: 15,
      })),
    ),
  },
  { title: 'Thực nhận', id: 'thucNhan', bodyStyle: 'money', width: 15 },
  { title: 'Ký nhận', id: 'kyNhan' },
] as const;

const workbook = defineWorkbook({
  defaultStyle: {
    font: { name: 'Times New Roman' },
  },
  styles: {
    money: {
      numFmt: '#,##0',
      alignment: { horizontal: 'right', vertical: 'middle' },
    },
    number: {
      numFmt: '#,##0',
      alignment: { horizontal: 'right', vertical: 'middle' },
    },
    title: {
      font: { bold: true, size: 14 },
      alignment: { horizontal: 'center', vertical: 'middle' },
    },
    header: {
      font: { bold: true, size: 11 },
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    },
  },
  context: {
    sheetTitle: 'DANH SÁCH GIẢNG VIÊN VƯỢT GIỜ NĂM 2024-2025',
  },
  sheets: [
    {
      id: 'summary',
      name: 'Tổng hợp',
      blocks: [
        {
          type: 'table',
          border: 'thin',
          headerStyle: 'header',
          headerRowHeights: [28, 28, 28],
          titleRows: [
            { value: '{{sheetTitle}}', style: 'title', height: 28 },
            { value: 'TỔNG HỢP TẤT CẢ CÁC KHOA', style: 'title', height: 24 },
            { value: '', height: 24 },
          ],
          columns,
          data: [],
        },
      ],
    },
    {
      id: 'cntt',
      name: 'Khoa CNTT',
      blocks: [
        {
          type: 'table',
          border: 'thin',
          headerStyle: 'header',
          headerRowHeights: [28, 28, 28],
          titleRows: [
            { value: '{{sheetTitle}}', style: 'title', height: 28 },
            { value: 'KHOA CÔNG NGHỆ THÔNG TIN', style: 'title', height: 24 },
            { value: '', height: 24 },
          ],
          columns,
          data: [
            {
              type: 'section',
              cells: [
                {
                  value: 'I',
                  style: { font: { bold: true }, alignment: { horizontal: 'center' } },
                },
                {
                  value: 'Khoa Công nghệ thông tin',
                  style: { font: { bold: true } },
                  colSpan: 'remaining',
                },
              ],
            },
            {
              stt: 1,
              giangVien: 'Nguyễn Văn A',
              thuNhap: 15000000,
              dinhMucGioGiang: 200,
              duocGiam: 20,
              soTietChuaHoanThanhNCKH: 10,
              dinhMucPhaiGiang: 170,
              hocKyI_vn: 80,
              hocKyI_lao: 10,
              hocKyI_cuba: 5,
              hocKyI_cpc: 0,
              hocKyI_dong_hp: 0,
              hocKyII_vn: 70,
              hocKyII_lao: 5,
              hocKyII_cuba: 5,
              hocKyII_cpc: 0,
              hocKyII_dong_hp: 0,
              caNam_vn: 150,
              caNam_lao: 15,
              caNam_cuba: 10,
              caNam_cpc: 0,
              caNam_dong_hp: 0,
              soTietVuotDinhMuc: 30,
              mucTTChuan: 100000,
              vn: 8000000,
              lao: 1500000,
              cuba: 1000000,
              cpc: 0,
              dong_hp: 0,
              tong: 10500000,
              thucNhan: {
                type: 'binary',
                operator: '-',
                left: { type: 'ref', id: 'tong' },
                right: { type: 'ref', id: 'mucTTChuan' },
              },
              kyNhan: '',
            },
            {
              type: 'section',
              style: { font: { bold: true }, alignment: { horizontal: 'right' } },
              cells: [
                {
                  value: 'Tổng',
                  style: { font: { bold: true }, alignment: { horizontal: 'center' } },
                  colSpan: 2,
                },
                {
                  id: 'tongSoTietChuaHoanThanhNCKH',
                  columnId: 'soTietChuaHoanThanhNCKH',
                  value: {
                    type: 'sum',
                    range: {
                      startId: 'soTietChuaHoanThanhNCKH',
                      endId: 'soTietChuaHoanThanhNCKH',
                      scope: 'currentRows',
                    },
                  },
                },
                {
                  id: 'tongHocKyI_vn',
                  columnId: 'hocKyI_vn',
                  value: {
                    type: 'sum',
                    range: {
                      startId: 'hocKyI_vn',
                      endId: 'hocKyI_vn',
                      scope: 'currentRows',
                    },
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: 'matma',
      name: 'Khoa Mật mã',
      blocks: [
        {
          type: 'table',
          border: 'thin',
          headerStyle: 'header',
          headerRowHeights: [28, 28, 28],
          titleRows: [
            { value: '{{sheetTitle}}', style: 'title', height: 28 },
            { value: 'KHOA MẬT MÃ', style: 'title', height: 24 },
            { value: '', height: 24 },
          ],
          columns,
          data: [],
        },
      ],
    },
  ],
});

export async function exportExcel(): Promise<void> {
  const outputPath = 'output/overtime-report-generated.xlsx';
  await mkdir('output', { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);
}

exportExcel()
  .then(() => {
    console.log('Excel exported successfully');
  })
  .catch((error) => {
    console.error('Error exporting Excel:', error);
  });
