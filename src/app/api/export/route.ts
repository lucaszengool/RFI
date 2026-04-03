import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { SUMMARY_HEADERS, type SupplierSummary } from '@/lib/extractors';

export async function POST(request: NextRequest) {
  try {
    const { data } = await request.json() as { data: SupplierSummary[] };

    if (!data || data.length === 0) {
      return NextResponse.json({ error: '没有数据可导出' }, { status: 400 });
    }

    const wb = XLSX.utils.book_new();

    // Category header row
    const catRow = ['', '', '', '', '', '', 'I平', '网络', '商务',
      '能评及资质', '', '', '', '', '',
      '供应及交付', '', '', '', '', '',
      '技术', '', '运营能力', '', '问题总结'];

    // Column headers
    const headers = SUMMARY_HEADERS.map(h => h.label);

    // Data rows
    const rows = data.map((item, idx) => {
      return SUMMARY_HEADERS.map(h => {
        if (h.key === 'id') return String(idx + 1);
        return (item as unknown as Record<string, string>)[h.key] || '';
      });
    });

    const wsData = [catRow, headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Set column widths
    ws['!cols'] = SUMMARY_HEADERS.map(h => ({ wch: Math.floor(h.width / 7) }));

    // Set merges for category headers
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // blank
      { s: { r: 0, c: 9 }, e: { r: 0, c: 14 } }, // 能评及资质
      { s: { r: 0, c: 21 }, e: { r: 0, c: 22 } }, // 技术
      { s: { r: 0, c: 23 }, e: { r: 0, c: 24 } }, // 运营能力
    ];

    XLSX.utils.book_append_sheet(wb, ws, '华东新园区');

    // Add 分级建议 sheet
    const classData = [
      ['腾讯大型数据中心选址分级标准'],
      ['级别', '标准描述（非现货机房、定制化租用）', '标准描述（现货机房）'],
      ['预A', '1、资源能力：能评及合规文件完整 ；针对全新园区（>100MW）场景，需在东数西算规定节点。\n2、技术能力：技术标准全部满足《技术综合评估表》；\n3、运营能力：建成并运营机架数量50000架（容量350MW）以上或属于现有合作同园区扩容；\n4、交付能力：供应及交付保障在业务指定日期内完成；', ''],
      ['预B', '1、资源能力：能评及合规文件完整 ；\n2、技术能力：技术标准全部满足《技术综合评估表》；\n3、运营能力：建成并运营机架数量10000架（容量70MW）~50000架（容量350MW）；\n4、交付能力：供应及交付保障有延期30天内风险；', ''],
      ['预C', '1、资源能力：能评及合规文件完整 ；\n2、技术能力：技术标准全部满足《技术综合评估表》；\n3、运营能力：建成并运营机架数量1000~10000架（容量7-70MW）；\n4、交付能力：供应及交付保障有延期60天内风险；', ''],
      ['预D', '1、资源能力：无能评及合规文件缺失；\n2、技术能力：技术标准无法满足；\n3、运营能力：建成并运营机架数量小于1000架（容量7MW）；\n4、交付能力：供应及交付保障有延期60天以上风险；', ''],
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(classData);
    ws2['!cols'] = [{ wch: 8 }, { wch: 60 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, ws2, '分级建议');

    // Add 现货场景关键品类清单 sheet
    const keyEquipData = [
      ['现货场景'],
      ['序号', '关键设备品类', '品牌'],
      ['1', '柴油发电机', '以《腾讯合建数据中心主要器件设备推荐选型名单》最新版本品牌为准'],
      ['2', '断路器及空气开关', ''],
      ['3', '蓄电池', ''],
      ['4', 'UPS', ''],
      ['5', 'HVDC', ''],
      ['6', '冷水机组', ''],
      ['7', '精密空调末端', ''],
      ['8', '水泵', ''],
      ['9', '电机', ''],
      ['10', '冷却塔', ''],
    ];
    const ws3 = XLSX.utils.aoa_to_sheet(keyEquipData);
    XLSX.utils.book_append_sheet(wb, ws3, '现货场景关键品类清单');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    return new NextResponse(buf, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="RFI_Summary_${new Date().toISOString().split('T')[0]}.xlsx"`,
      },
    });
  } catch (error) {
    console.error('Export error:', error);
    return NextResponse.json({ error: '导出失败' }, { status: 500 });
  }
}
