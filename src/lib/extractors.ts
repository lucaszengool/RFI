import * as XLSX from 'xlsx';

export interface SupplierSummary {
  id: string;
  type: string; // 类型
  city: string; // 城市
  dcName: string; // 机房名称
  supplier: string; // 供应商
  address: string; // 地址
  rating: string; // 评级
  credit: string; // 信用
  dangerZone: string; // 是否处于危险地带
  energyEvalCoal: string; // 能评(煤当量值)
  energyEvalIT: string; // 能评量(IT MW)
  currentForTencent: string; // 当前给腾讯(MW)
  energyExpansion: string; // 能评扩容(MW)
  afterExpansion: string; // 扩容后(MW)
  idcLicense: string; // IDC牌照
  landStatus: string; // 土地情况
  currentLand: string; // 现有土地(亩)
  landExpansion: string; // 土地扩容
  afterLandExpansion: string; // 扩容后(亩)
  powerSupply: string; // 供电
  deliverySchedule: string; // 交付工期
  architecturePlan: string; // 架构方案
  whitelist: string; // 白名单
  operationScale: string; // 运营规模
  service: string; // 服务
  issueSummary: string; // 问题总结
}

// Column header mapping for the summary table
export const SUMMARY_HEADERS = [
  { key: 'id', label: '序号', width: 60 },
  { key: 'type', label: '类型', width: 80 },
  { key: 'city', label: '城市', width: 80 },
  { key: 'dcName', label: '机房名称', width: 200 },
  { key: 'supplier', label: '供应商', width: 180 },
  { key: 'address', label: '地址', width: 250 },
  { key: 'rating', label: '评级', width: 60 },
  { key: 'credit', label: '信用', width: 80 },
  { key: 'dangerZone', label: '是否处于危险地带', width: 120 },
  { key: 'energyEvalCoal', label: '能评\n（煤当量值）', width: 120 },
  { key: 'energyEvalIT', label: '能评量\n（IT MW）', width: 100 },
  { key: 'currentForTencent', label: '当前给腾讯\n（MW）', width: 100 },
  { key: 'energyExpansion', label: '能评扩容（MW）', width: 100 },
  { key: 'afterExpansion', label: '扩容后（MW）', width: 100 },
  { key: 'idcLicense', label: 'IDC牌照', width: 80 },
  { key: 'landStatus', label: '土地情况', width: 200 },
  { key: 'currentLand', label: '现有土地（亩）', width: 100 },
  { key: 'landExpansion', label: '土地扩容', width: 100 },
  { key: 'afterLandExpansion', label: '扩容后（亩）', width: 100 },
  { key: 'powerSupply', label: '供电', width: 300 },
  { key: 'deliverySchedule', label: '交付工期', width: 250 },
  { key: 'architecturePlan', label: '架构方案', width: 250 },
  { key: 'whitelist', label: '白名单', width: 200 },
  { key: 'operationScale', label: '运营规模', width: 200 },
  { key: 'service', label: '服务', width: 200 },
  { key: 'issueSummary', label: '问题总结', width: 300 },
];

export const CATEGORY_HEADERS = [
  { label: '', span: 1 }, // 序号
  { label: '', span: 1 }, // 类型
  { label: '', span: 1 }, // 城市
  { label: '', span: 1 }, // 机房名称
  { label: '', span: 1 }, // 供应商
  { label: '', span: 1 }, // 地址
  { label: 'I平', span: 1 }, // 评级
  { label: '网络', span: 1 }, // (placeholder)
  { label: '商务', span: 1 }, // 信用
  { label: '能评及资质', span: 6 }, // 能评...IDC牌照
  { label: '供应及交付', span: 1 }, // 土地情况 onwards
  { label: '技术', span: 2 }, // 架构方案+白名单
  { label: '运营能力', span: 2 }, // 运营规模+服务
  { label: '问题总结', span: 1 },
];

function getCellValue(ws: XLSX.WorkSheet, cell: string): string {
  const c = ws[cell];
  if (!c) return '';
  return String(c.v ?? '').trim();
}

function findCellByContent(ws: XLSX.WorkSheet, searchText: string): string | null {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && String(cell.v ?? '').includes(searchText)) {
        return addr;
      }
    }
  }
  return null;
}

function findValueAfterLabel(ws: XLSX.WorkSheet, label: string, colOffset = 1): string {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && String(cell.v ?? '').includes(label)) {
        // Look for value in cells to the right
        for (let offset = colOffset; offset <= colOffset + 3; offset++) {
          const valAddr = XLSX.utils.encode_cell({ r, c: c + offset });
          const valCell = ws[valAddr];
          if (valCell && String(valCell.v ?? '').trim()) {
            return String(valCell.v ?? '').trim();
          }
        }
      }
    }
  }
  return '';
}

function getSupplierResponseCol(ws: XLSX.WorkSheet): number {
  // Find the column labeled "供应商应答" or "供应商填写"
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let r = 0; r <= Math.min(range.e.r, 3); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && (String(cell.v ?? '').includes('供应商应答') || String(cell.v ?? '').includes('供应商填写'))) {
        return c;
      }
    }
  }
  return -1;
}

function getResponseForItem(ws: XLSX.WorkSheet, itemText: string, responseCol: number): string {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && String(cell.v ?? '').includes(itemText)) {
        const valAddr = XLSX.utils.encode_cell({ r, c: responseCol });
        const valCell = ws[valAddr];
        if (valCell) return String(valCell.v ?? '').trim();
      }
    }
  }
  return '';
}

export function extractSupplierData(workbook: XLSX.WorkBook): SupplierSummary {
  const summary: SupplierSummary = {
    id: '',
    type: '主园区',
    city: '',
    dcName: '',
    supplier: '',
    address: '',
    rating: '',
    credit: '',
    dangerZone: '',
    energyEvalCoal: '',
    energyEvalIT: '',
    currentForTencent: '',
    energyExpansion: '',
    afterExpansion: '',
    idcLicense: '',
    landStatus: '',
    currentLand: '',
    landExpansion: '',
    afterLandExpansion: '',
    powerSupply: '',
    deliverySchedule: '',
    architecturePlan: '',
    whitelist: '',
    operationScale: '',
    service: '',
    issueSummary: '',
  };

  // Sheet 1: 供应商、土地、市电、机柜等信息
  const sheet1Name = workbook.SheetNames.find(n => n.includes('供应商') || n.includes('土地'));
  if (sheet1Name) {
    const ws1 = workbook.Sheets[sheet1Name];

    // Basic info
    summary.dcName = findValueAfterLabel(ws1, '名称', 1);

    // Address
    const addrVal = findValueAfterLabel(ws1, '地址', 1);
    summary.address = addrVal.split('\n')[0].replace(/^a\.\s*/, '').trim();

    // City extraction from address
    if (summary.address.includes('北京')) summary.city = '北京';
    else if (summary.address.includes('上海')) summary.city = '上海';
    else if (summary.address.includes('广州')) summary.city = '广州';
    else if (summary.address.includes('深圳')) summary.city = '深圳';
    else {
      const cityMatch = summary.address.match(/^(.{2,3})[市省]/);
      if (cityMatch) summary.city = cityMatch[1];
    }

    // Danger zone
    summary.dangerZone = findValueAfterLabel(ws1, '山地、倾斜地') ||
                          findValueAfterLabel(ws1, '危险地带') || '否';

    // Land info
    const landArea = findValueAfterLabel(ws1, '建设用地面积');
    const totalBuildingArea = findValueAfterLabel(ws1, '总建筑面积');

    // Calculate land in 亩 (1 m² = 0.0015 亩)
    const landMatch = landArea.match(/([\d.]+)/);
    if (landMatch) {
      const sqm = parseFloat(landMatch[1]);
      const mu = (sqm / 666.67).toFixed(1);
      summary.currentLand = mu;
      summary.afterLandExpansion = mu;
    }

    summary.landStatus = `总用地规模：${summary.currentLand}亩\n总建筑面积：${totalBuildingArea}`;

    // Power supply - extract substation info (PPT Slide 3 requirement)
    const substationNames: string[] = [];
    const upperStation = findValueAfterLabel(ws1, '上一级已经投产变电站', 2);
    if (upperStation) substationNames.push(upperStation);

    const localStation = findValueAfterLabel(ws1, '负责本园区', 2);
    const localCapacity = findValueAfterLabel(ws1, '规划容量', 1);

    summary.powerSupply = '';
    if (localStation) {
      summary.powerSupply += `1、负责本园区：${localStation}`;
      if (localCapacity) summary.powerSupply += `（${localCapacity}）`;
    }
    if (upperStation) {
      summary.powerSupply += `\n2、上级变电站${upperStation}`;
    }

    // Delivery schedule and architecture (PPT Slide 4 - buildings, capacity, cooling, timeline)
    const futureInfo = findValueAfterLabel(ws1, '期货情况', 1);
    const futureType = findValueAfterLabel(ws1, '期货情况', 2);
    const futureTime = findValueAfterLabel(ws1, '期货情况', 3);

    let deliveryParts: string[] = [];
    if (futureTime) deliveryParts.push(`1、${futureTime}`);
    if (futureInfo) {
      // Combine building info
      const buildings: string[] = [];
      const range = XLSX.utils.decode_range(ws1['!ref'] || 'A1');
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r, c });
          const cell = ws1[addr];
          if (cell && String(cell.v ?? '').includes('#楼')) {
            buildings.push(String(cell.v ?? '').trim());
          }
        }
      }
    }
    if (futureType) deliveryParts.push(`2、园区规划${futureType}`);
    summary.deliverySchedule = deliveryParts.join('\n') || futureTime || '';

    // Electricity info
    const powerNature = findValueAfterLabel(ws1, '机房用电性质', 1);
    if (powerNature && !summary.powerSupply.includes(powerNature)) {
      // Already captured above
    }
  }

  // Sheet 2: 技术综合评估表
  const sheet2Name = workbook.SheetNames.find(n => n.includes('技术综合评估'));
  if (sheet2Name) {
    const ws2 = workbook.Sheets[sheet2Name];
    const respCol = getSupplierResponseCol(ws2);

    // IDC License
    const idcResp = getResponseForItem(ws2, 'IDC牌照', respCol >= 0 ? respCol : 6);
    summary.idcLicense = idcResp.includes('满足') ? '有' : idcResp || '';

    // Energy evaluation (能评)
    const energyResp = getResponseForItem(ws2, '能评数据', respCol >= 0 ? respCol : 6);
    // Extract coal equivalent
    const coalMatch = energyResp.match(/当量值[：:]\s*([\d,]+)/);
    if (coalMatch) {
      summary.energyEvalCoal = coalMatch[1].replace(/,/g, '');
    } else {
      const coalMatch2 = energyResp.match(/([\d,]+)\s*吨/);
      if (coalMatch2) summary.energyEvalCoal = coalMatch2[1].replace(/,/g, '');
    }

    // Extract IT MW
    const itMWMatch = energyResp.match(/IT产出\s*([\d.]+)\s*MW/i);
    if (itMWMatch) {
      summary.energyEvalIT = itMWMatch[1];
      summary.currentForTencent = itMWMatch[1];
    }

    // Energy expansion
    const expansionResp = getResponseForItem(ws2, '扩容需求能评', respCol >= 0 ? respCol : 6);
    const expMatch = expansionResp.match(/([\d.]+)\+?\s*MW/i);
    if (expMatch) {
      summary.energyExpansion = String(parseFloat(expMatch[1]) - parseFloat(summary.energyEvalIT || '0'));
    } else if (expansionResp.includes('100')) {
      // If mentions 100MW expansion
      summary.energyExpansion = String(100 - parseFloat(summary.energyEvalIT || '0'));
      summary.afterExpansion = String(100 + parseFloat(summary.energyEvalIT || '0'));
    }

    if (summary.energyEvalIT && summary.energyExpansion) {
      summary.afterExpansion = String(parseFloat(summary.energyEvalIT) + parseFloat(summary.energyExpansion));
    }

    // Supplier name from company info
    const scaleResp = getResponseForItem(ws2, '贵司整体投运', respCol >= 0 ? respCol : 6);
    if (scaleResp) {
      const nameMatch = scaleResp.match(/^([^\s截]+)/);
      if (nameMatch) {
        // Try to get supplier from the scale response
      }
    }

    // Operation scale (PPT Slide 6)
    summary.operationScale = '';
    if (scaleResp) {
      // Extract total IT capacity
      const capMatch = scaleResp.match(/([\d,]+)\s*MW/);
      if (capMatch) {
        summary.operationScale = `截止目前的IT负载容量为\n约${capMatch[1]}MW（包含规划在建）`;
      }
    }

    // Concurrent delivery
    const concurrentResp = getResponseForItem(ws2, '并发交付能力', respCol >= 0 ? respCol : 6);

    // SLA / Service (PPT Slide 6)
    // Check operation-related responses
    const slaItems = ['维保策略', '日常巡检', '应急演练', '变更管理', '事件管理', '备件管理', '联合运营'];
    let allMet = true;
    const unmetItems: string[] = [];

    for (const item of slaItems) {
      const resp = getResponseForItem(ws2, item, respCol >= 0 ? respCol : 6);
      if (resp && !resp.includes('满足')) {
        allMet = false;
        unmetItems.push(item);
      }
    }

    if (allMet) {
      summary.service = '1、满足SLA 2.4；\n2、重视程度，配合意愿，资源投入和能力较好';
    } else {
      summary.service = `不满足项：${unmetItems.join('、')}`;
    }

    // Architecture plan (PPT Slide 4 - cooling type)
    const coolingResp = getResponseForItem(ws2, '建筑设计要求', respCol >= 0 ? respCol : 6);

    // Whitelist (PPT Slide 5)
    const whitelistItems = ['UPS', 'HVDC', '柴油发电机', '电池', 'AHU', 'PHU', 'SHU', 'MAC', '弱电', '液冷'];
    const nonCompliant: string[] = [];
    let allWhitelist = true;

    for (const item of whitelistItems) {
      const resp = getResponseForItem(ws2, item, respCol >= 0 ? respCol : 6);
      if (resp && !resp.includes('满足') && resp !== '') {
        allWhitelist = false;
        nonCompliant.push(`${item}: ${resp}`);
      }
    }

    summary.whitelist = allWhitelist ? '非现货机房，设备可参照白名单定制' :
      `不符合白名单设备：\n${nonCompliant.join('\n')}`;

    // Architecture - extract cooling info
    const archParts: string[] = [];
    // Check for cooling type mentions
    const range2 = XLSX.utils.decode_range(ws2['!ref'] || 'A1');
    for (let r = range2.s.r; r <= range2.e.r; r++) {
      const gAddr = XLSX.utils.encode_cell({ r, c: respCol >= 0 ? respCol : 6 });
      const gCell = ws2[gAddr];
      if (gCell) {
        const val = String(gCell.v ?? '');
        if (val.includes('冷源') || val.includes('冷冻水') || val.includes('液冷')) {
          if (!archParts.includes(val.trim())) {
            archParts.push(val.trim());
          }
        }
      }
    }

    if (archParts.length > 0) {
      summary.architecturePlan = archParts.join('\n');
    }

    // Issue summary (PPT Slide 7)
    // Look for items that are NOT "满足"
    const issues: string[] = [];

    // Check energy eval validity
    if (energyResp.includes('2019') || energyResp.includes('2020')) {
      issues.push('能评报告有效期问题');
    }

    // Check for specific issues mentioned in responses
    for (let r = range2.s.r; r <= range2.e.r; r++) {
      const gAddr = XLSX.utils.encode_cell({ r, c: respCol >= 0 ? respCol : 6 });
      const gCell = ws2[gAddr];
      if (gCell) {
        const val = String(gCell.v ?? '');
        if (val.includes('偏离') || val.includes('不满足') || val.includes('暂无') || val.includes('不涉及')) {
          // Get the requirement context
          for (let c = 0; c <= 4; c++) {
            const ctxAddr = XLSX.utils.encode_cell({ r, c });
            const ctxCell = ws2[ctxAddr];
            if (ctxCell && String(ctxCell.v ?? '').trim()) {
              issues.push(`${String(ctxCell.v ?? '').trim()}：${val.substring(0, 50)}`);
              break;
            }
          }
        }
      }
    }

    summary.issueSummary = issues.length > 0 ? issues.join('\n') : '无';
  }

  // Try to extract supplier name from filename or sheet data
  if (!summary.supplier) {
    // Look in first sheet for company info
    const sheet1Name = workbook.SheetNames.find(n => n.includes('供应商') || n.includes('土地'));
    if (sheet1Name) {
      const ws1 = workbook.Sheets[sheet1Name];
      const fillPerson = findValueAfterLabel(ws1, '填表人');
      // Try to find company name from the workbook
      summary.supplier = findValueAfterLabel(ws1, '名称', 1) ?
        summary.dcName.replace(/数据中心.*/, '').replace(/机房.*/, '') : '';
    }
  }

  return summary;
}

export function exportToExcel(data: SupplierSummary[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Create headers row
  const headers = SUMMARY_HEADERS.map(h => h.label);
  const rows = data.map((item, idx) => {
    return SUMMARY_HEADERS.map(h => {
      if (h.key === 'id') return String(idx + 1);
      return (item as unknown as Record<string, string>)[h.key] || '';
    });
  });

  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Set column widths
  ws['!cols'] = SUMMARY_HEADERS.map(h => ({ wch: Math.floor(h.width / 8) }));

  XLSX.utils.book_append_sheet(wb, ws, '华东新园区');

  // Add classification sheet
  const classData = [
    ['腾讯大型数据中心选址分级标准'],
    ['级别', '标准描述（非现货机房、定制化租用）', '标准描述（现货机房）'],
    ['预A', '1、资源能力：能评及合规文件完整；2、技术标准全部满足；3、运营机架>50000架(>350MW)；4、交付按期完成', ''],
    ['预B', '1、资源能力：能评及合规文件完整；2、技术标准全部满足；3、运营机架10000-50000架(70-350MW)；4、交付延期<30天', ''],
    ['预C', '1、资源能力：能评及合规文件完整；2、技术标准全部满足；3、运营机架1000-10000架(7-70MW)；4、交付延期<60天', ''],
    ['预D', '1、无能评/合规文件缺失；2、技术标准无法满足；3、运营机架<1000架(<7MW)；4、交付延期>60天', ''],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(classData);
  XLSX.utils.book_append_sheet(wb, ws2, '分级建议');

  return wb;
}
