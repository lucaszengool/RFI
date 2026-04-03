import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { extractSupplierData } from '@/lib/extractors';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return NextResponse.json({ error: '请上传文件' }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const workbook = XLSX.read(buffer, { type: 'buffer' });

    const extracted = extractSupplierData(workbook);

    // Use filename to enrich data
    const fileName = file.name;
    if (fileName.includes('北京')) extracted.city = extracted.city || '北京';
    if (fileName.includes('上海')) extracted.city = extracted.city || '上海';
    if (fileName.includes('普洛斯')) extracted.supplier = extracted.supplier || '杭州普璋数据科技有限公司';

    return NextResponse.json({
      success: true,
      data: extracted,
      fileName: file.name,
      sheets: workbook.SheetNames,
    });
  } catch (error) {
    console.error('Upload error:', error);
    return NextResponse.json(
      { error: '文件解析失败，请确认上传的是供应商评估表Excel文件' },
      { status: 500 }
    );
  }
}

