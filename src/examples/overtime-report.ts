import { mkdir } from 'node:fs/promises';
import { renderWorkbook } from '../compile';
import { textBlock } from '../factories';
import { f } from '../formula';
import type { WorkbookDefinition } from '../types';

const danhSachDoiTuong = [
  { title: 'VN', id: 'vn' },
  { title: 'Lào', id: 'lao' },
  { title: 'Cuba', id: 'cuba' },
  { title: 'CPC', id: 'cpc' },
  { title: 'Đóng HP', id: 'dong_hp' },
];

const columns = [
  { title: 'STT', id: 'stt', bodyStyle: { alignment: { horizontal: 'center' as const } } },
  {
    title: 'Họ tên giảng viên',
    id: 'giangVien',
    bodyStyle: { alignment: { horizontal: 'left' as const, wrapText: true } },
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
        title: doiTuong.title,
        id: doiTuong.id,
        bodyStyle: 'money',
        width: 15,
      })),
    ),
  },
  { title: 'Thực nhận', id: 'thucNhan', bodyStyle: 'money', width: 15 },
  { title: 'Ký nhận', id: 'kyNhan' },
] as const;

function createFacultyFooter(facultyId: string) {
  const numberColumns = [
    'thuNhap',
    'dinhMucGioGiang',
    'duocGiam',
    'soTietChuaHoanThanhNCKH',
    'dinhMucPhaiGiang',
    ...danhSachDoiTuong.map((d) => `hocKyI_${d.id}`),
    ...danhSachDoiTuong.map((d) => `hocKyII_${d.id}`),
    ...[...danhSachDoiTuong, { id: 'tong' }].map((d) => `caNam_${d.id}`),
    'soTietVuotDinhMuc',
    ...[...danhSachDoiTuong, { id: 'tong' }].map((d) => d.id),
    'thucNhan',
  ];

  return {
    style: { font: { bold: true }, alignment: { horizontal: 'right' as const } },
    cells: [
      {
        value: 'Tổng',
        style: { font: { bold: true }, alignment: { horizontal: 'center' as const } },
        colSpan: 2,
      },
      ...numberColumns.map((colId) => ({
        id: `${facultyId}_${colId}`,
        columnId: colId,
        value: f`SUM(${f.range(colId, colId, { scope: 'currentRows' })})`,
      })),
    ],
  };
}

function getExcelColumn(colIndex: number): string {
  let temp = colIndex;
  let letter = '';
  while (temp > 0) {
    temp--;
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26);
  }
  return letter;
}

const leafColumnIds = [
  'stt',
  'giangVien',
  'thuNhap',
  'dinhMucGioGiang',
  'duocGiam',
  'soTietChuaHoanThanhNCKH',
  'dinhMucPhaiGiang',
  ...danhSachDoiTuong.map((d) => `hocKyI_${d.id}`),
  ...danhSachDoiTuong.map((d) => `hocKyII_${d.id}`),
  ...[...danhSachDoiTuong, { id: 'tong' }].map((d) => `caNam_${d.id}`),
  'soTietVuotDinhMuc',
  'mucTTChuan',
  ...[...danhSachDoiTuong, { id: 'tong' }].map((d) => d.id),
  'thucNhan',
  'kyNhan',
];

function createSummaryDataForFaculty(facultyName: string, rowCount: number, startRowOffset: number) {
  const data = [];
  for (let i = 0; i < rowCount; i++) {
    const targetRow = startRowOffset + i;
    const rowData: Record<string, any> = {};
    leafColumnIds.forEach((colId, colIndex) => {
      const colLetter = getExcelColumn(colIndex + 1);
      // Raw formula linking to the faculty sheet
      rowData[colId] = f`'${facultyName}'!${colLetter}${targetRow}`;
    });
    data.push(rowData);
  }
  return data;
}

function createSummaryFooter() {
  const numberColumns = [
    'thuNhap',
    'dinhMucGioGiang',
    'duocGiam',
    'soTietChuaHoanThanhNCKH',
    'dinhMucPhaiGiang',
    ...danhSachDoiTuong.map((d) => `hocKyI_${d.id}`),
    ...danhSachDoiTuong.map((d) => `hocKyII_${d.id}`),
    ...[...danhSachDoiTuong, { id: 'tong' }].map((d) => `caNam_${d.id}`),
    'soTietVuotDinhMuc',
    ...[...danhSachDoiTuong, { id: 'tong' }].map((d) => d.id),
    'thucNhan',
  ];

  return {
    style: { font: { bold: true }, alignment: { horizontal: 'right' as const } },
    cells: [
      {
        value: 'Tổng toàn trường',
        style: { font: { bold: true }, alignment: { horizontal: 'center' as const } },
        colSpan: 2,
      },
      ...numberColumns.map((colId) => ({
        columnId: colId,
        value: f`SUM(${f.range(colId, colId, { scope: 'allRows' })})`,
      })),
    ],
  };
}

function createFacultySheet(id: string, name: string, facultyTitle: string, data: any[]) {
  return {
    id,
    name,
    blocks: [
      textBlock('{{sheetTitle}}', { style: 'title', height: 28, colSpan: 'remaining' }),
      textBlock(facultyTitle, { style: 'title', height: 24, colSpan: 'remaining' }),
      textBlock('', { height: 24, colSpan: 'remaining' }),
      {
        type: 'table' as const,
        border: 'thin' as const,
        headerStyle: 'header',
        headerRowHeights: [28, 28, 28],
        columns,
        data,
        footerRows: [createFacultyFooter(id)],
      },
    ],
  };
}

const workbook: WorkbookDefinition = {
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
        textBlock('{{sheetTitle}}', { style: 'title', height: 28, colSpan: 'remaining' }),
        textBlock('TỔNG HỢP TẤT CẢ CÁC KHOA', { style: 'title', height: 24, colSpan: 'remaining' }),
        textBlock('', { height: 24, colSpan: 'remaining' }),
        {
          type: 'table' as const,
          border: 'thin' as const,
          headerStyle: 'header',
          headerRowHeights: [28, 28, 28],
          columns,
          groups: [
            {
              headerRows: [
                {
                  cells: [
                    { value: 'I', style: { font: { bold: true }, alignment: { horizontal: 'center' as const } } },
                    {
                      value: 'Khoa Công nghệ thông tin',
                      style: { font: { bold: true } },
                      colSpan: 'remaining' as const,
                    },
                  ],
                },
              ],
              data: createSummaryDataForFaculty('Khoa CNTT', 2, 7),
              footerRows: [createFacultyFooter('cntt')],
            },
            {
              headerRows: [
                {
                  cells: [
                    { value: 'II', style: { font: { bold: true }, alignment: { horizontal: 'center' as const } } },
                    { value: 'Khoa Mật mã', style: { font: { bold: true } }, colSpan: 'remaining' as const },
                  ],
                },
              ],
              data: createSummaryDataForFaculty('Khoa Mật mã', 1, 7),
              footerRows: [createFacultyFooter('matma')],
            },
          ],
          footerRows: [createSummaryFooter()],
        },
      ],
    },
    createFacultySheet('cntt', 'Khoa CNTT', 'KHOA CÔNG NGHỆ THÔNG TIN', [
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
        caNam_tong: 175,
        soTietVuotDinhMuc: 30,
        mucTTChuan: 100000,
        vn: 8000000,
        lao: 1500000,
        cuba: 1000000,
        cpc: 0,
        dong_hp: 0,
        tong: 10500000,
        thucNhan: f`${f.ref('tong')} - ${f.ref('mucTTChuan')}`,
        kyNhan: '',
      },
      {
        stt: 2,
        giangVien: 'Trần Thị B',
        thuNhap: 10000000,
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
        caNam_tong: 175,
        soTietVuotDinhMuc: 30,
        mucTTChuan: 100000,
        vn: 8000000,
        lao: 1500000,
        cuba: 1000000,
        cpc: 0,
        dong_hp: 0,
        tong: 10500000,
        thucNhan: f`${f.ref('tong')} - ${f.ref('mucTTChuan')}`,
        kyNhan: '',
      },
    ]),
    createFacultySheet('matma', 'Khoa Mật mã', 'KHOA MẬT MÃ', [
      {
        stt: 1,
        giangVien: 'Lê Văn C',
        thuNhap: 12000000,
        dinhMucGioGiang: 200,
        duocGiam: 0,
        soTietChuaHoanThanhNCKH: 0,
        dinhMucPhaiGiang: 200,
        hocKyI_vn: 100,
        hocKyI_lao: 0,
        hocKyI_cuba: 0,
        hocKyI_cpc: 0,
        hocKyI_dong_hp: 0,
        hocKyII_vn: 150,
        hocKyII_lao: 0,
        hocKyII_cuba: 0,
        hocKyII_cpc: 0,
        hocKyII_dong_hp: 0,
        caNam_vn: 250,
        caNam_lao: 0,
        caNam_cuba: 0,
        caNam_cpc: 0,
        caNam_dong_hp: 0,
        caNam_tong: 250,
        soTietVuotDinhMuc: 50,
        mucTTChuan: 100000,
        vn: 5000000,
        lao: 0,
        cuba: 0,
        cpc: 0,
        dong_hp: 0,
        tong: 5000000,
        thucNhan: f`${f.ref('tong')} - ${f.ref('mucTTChuan')}`,
        kyNhan: '',
      },
    ]),
  ],
};

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
