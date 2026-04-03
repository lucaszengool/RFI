'use client';

import { useState, useCallback, useRef } from 'react';
import { SUMMARY_HEADERS, type SupplierSummary } from '@/lib/extractors';

const EMPTY_SUPPLIER: SupplierSummary = {
  id: '', type: '主园区', city: '', dcName: '', supplier: '', address: '',
  rating: '', credit: '', dangerZone: '', energyEvalCoal: '', energyEvalIT: '',
  currentForTencent: '', energyExpansion: '', afterExpansion: '', idcLicense: '',
  landStatus: '', currentLand: '', landExpansion: '', afterLandExpansion: '',
  powerSupply: '', deliverySchedule: '', architecturePlan: '', whitelist: '',
  operationScale: '', service: '', issueSummary: '',
};

type TabId = 'upload' | 'table' | 'classification';

export default function Home() {
  const [suppliers, setSuppliers] = useState<SupplierSummary[]>([]);
  const [activeTab, setActiveTab] = useState<TabId>('upload');
  const [uploading, setUploading] = useState(false);
  const [uploadLog, setUploadLog] = useState<string[]>([]);
  const [editingCell, setEditingCell] = useState<{ row: number; key: string } | null>(null);
  const [dragActive, setDragActive] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleUpload = useCallback(async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    setUploading(true);

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      setUploadLog(prev => [...prev, `正在解析: ${file.name}...`]);

      const formData = new FormData();
      formData.append('file', file);

      try {
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        const result = await res.json();

        if (result.success) {
          setSuppliers(prev => [...prev, { ...result.data, id: String(prev.length + 1) }]);
          setUploadLog(prev => [...prev, `[OK] ${file.name} 解析成功 - ${result.data.dcName || '未识别名称'}`]);
        } else {
          setUploadLog(prev => [...prev, `[FAIL] ${file.name} 解析失败: ${result.error}`]);
        }
      } catch (err) {
        setUploadLog(prev => [...prev, `[FAIL] ${file.name} 上传失败: ${String(err)}`]);
      }
    }

    setUploading(false);
    if (files.length > 0) setActiveTab('table');
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragActive(false);
    handleUpload(e.dataTransfer.files);
  }, [handleUpload]);

  const handleCellEdit = (rowIdx: number, key: string, value: string) => {
    setSuppliers(prev => {
      const updated = [...prev];
      updated[rowIdx] = { ...updated[rowIdx], [key]: value };
      return updated;
    });
  };

  const handleExport = async () => {
    try {
      const res = await fetch('/api/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: suppliers }),
      });
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `RFI_Summary_${new Date().toISOString().split('T')[0]}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert('导出失败: ' + String(err));
    }
  };

  const addEmptyRow = () => {
    setSuppliers(prev => [...prev, { ...EMPTY_SUPPLIER, id: String(prev.length + 1) }]);
  };

  const deleteRow = (idx: number) => {
    setSuppliers(prev => prev.filter((_, i) => i !== idx).map((s, i) => ({ ...s, id: String(i + 1) })));
  };

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100">
      {/* Header */}
      <header className="border-b border-gray-800 bg-gray-900/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-[1920px] mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center font-bold text-sm">
              RFI
            </div>
            <div>
              <h1 className="text-lg font-semibold">供应商RFI信息归整系统</h1>
              <p className="text-xs text-gray-500">腾讯大型数据中心选址评估</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs text-gray-500">
              已录入 {suppliers.length} 个供应商
            </span>
            {suppliers.length > 0 && (
              <button
                onClick={handleExport}
                className="px-4 py-2 bg-green-600 hover:bg-green-700 rounded-lg text-sm font-medium transition-colors"
              >
                导出Excel
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Tabs */}
      <div className="border-b border-gray-800 bg-gray-900/50">
        <div className="max-w-[1920px] mx-auto px-4 flex gap-1">
          {([
            { id: 'upload' as TabId, label: '上传解析' },
            { id: 'table' as TabId, label: '汇总表格' },
            { id: 'classification' as TabId, label: '分级标准' },
          ]).map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
                activeTab === tab.id
                  ? 'border-blue-500 text-blue-400'
                  : 'border-transparent text-gray-500 hover:text-gray-300'
              }`}
            >
              {tab.label}
              {tab.id === 'table' && suppliers.length > 0 && (
                <span className="ml-2 px-1.5 py-0.5 bg-blue-600/30 text-blue-400 rounded text-xs">
                  {suppliers.length}
                </span>
              )}
            </button>
          ))}
        </div>
      </div>

      {/* Content */}
      <main className="max-w-[1920px] mx-auto p-4">
        {/* Upload Tab */}
        {activeTab === 'upload' && (
          <div className="max-w-3xl mx-auto space-y-6">
            {/* Drop Zone */}
            <div
              onDragOver={(e) => { e.preventDefault(); setDragActive(true); }}
              onDragLeave={() => setDragActive(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
              className={`border-2 border-dashed rounded-xl p-12 text-center cursor-pointer transition-all ${
                dragActive
                  ? 'border-blue-500 bg-blue-500/10'
                  : 'border-gray-700 hover:border-gray-600 bg-gray-900/50'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                multiple
                onChange={(e) => handleUpload(e.target.files)}
                className="hidden"
              />
              <p className="text-4xl mb-4">
                {uploading ? '...' : ''}
              </p>
              <p className="text-lg font-medium mb-2">
                {uploading ? '正在解析文件...' : '拖拽或点击上传供应商评估表'}
              </p>
              <p className="text-sm text-gray-500">
                支持 .xlsx 格式，可同时上传多个文件
              </p>
              <p className="text-xs text-gray-600 mt-2">
                文件格式：腾讯租赁数据中心评估表V1.2
              </p>
            </div>

            {/* Upload Log */}
            {uploadLog.length > 0 && (
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-4">
                <h3 className="text-sm font-medium text-gray-400 mb-3">解析日志</h3>
                <div className="space-y-1 max-h-60 overflow-y-auto">
                  {uploadLog.map((log, i) => (
                    <div
                      key={i}
                      className={`text-sm font-mono ${
                        log.startsWith('[OK]') ? 'text-green-400' :
                        log.startsWith('[FAIL]') ? 'text-red-400' : 'text-gray-400'
                      }`}
                    >
                      {log}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Instructions */}
            <div className="bg-gray-900 border border-gray-800 rounded-xl p-6">
              <h3 className="text-sm font-semibold text-gray-300 mb-4">使用说明</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm text-gray-400">
                <div className="space-y-3">
                  <div>
                    <p className="font-medium text-gray-300 mb-1">1. 上传评估表</p>
                    <p>上传供应商填写的《腾讯租赁数据中心评估表》Excel文件</p>
                  </div>
                  <div>
                    <p className="font-medium text-gray-300 mb-1">2. 自动提取</p>
                    <p>系统自动提取关键信息：能评、供电、交付、白名单、运营等</p>
                  </div>
                </div>
                <div className="space-y-3">
                  <div>
                    <p className="font-medium text-gray-300 mb-1">3. 人工复核</p>
                    <p>在汇总表格中查看并编辑提取结果，所有单元格可直接编辑</p>
                  </div>
                  <div>
                    <p className="font-medium text-gray-300 mb-1">4. 导出汇总</p>
                    <p>点击"导出Excel"生成完整的RFI首轮信息归整表</p>
                  </div>
                </div>
              </div>

              <div className="mt-4 pt-4 border-t border-gray-800">
                <h4 className="text-xs font-semibold text-gray-500 mb-2">提取规则（基于PPT要求）</h4>
                <ul className="text-xs text-gray-600 space-y-1">
                  <li>- 供电：提取上级变电站数量、名称及对应容量</li>
                  <li>- 交付：涉及楼号、IT产出容量、制冷方案类型及可交付时间</li>
                  <li>- 白名单：标注是否所有设备符合腾讯白名单，列出不符合项</li>
                  <li>- 运营：历史运营规模、并发交付能力及SLA满足情况</li>
                  <li>- 问题：汇总不满足项</li>
                </ul>
              </div>
            </div>
          </div>
        )}

        {/* Summary Table Tab */}
        {activeTab === 'table' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <h2 className="text-lg font-semibold">供应商信息汇总</h2>
              <div className="flex gap-2">
                <button
                  onClick={addEmptyRow}
                  className="px-3 py-1.5 bg-gray-800 hover:bg-gray-700 rounded-lg text-sm transition-colors"
                >
                  + 添加行
                </button>
                <button
                  onClick={handleExport}
                  disabled={suppliers.length === 0}
                  className="px-3 py-1.5 bg-green-600 hover:bg-green-700 disabled:opacity-50 rounded-lg text-sm transition-colors"
                >
                  导出Excel
                </button>
              </div>
            </div>

            {suppliers.length === 0 ? (
              <div className="text-center py-20 text-gray-600">
                <p className="text-lg mb-2">暂无数据</p>
                <p className="text-sm">请先上传供应商评估表或手动添加行</p>
              </div>
            ) : (
              <div className="overflow-x-auto border border-gray-800 rounded-xl">
                <table className="w-full text-xs">
                  {/* Category headers */}
                  <thead>
                    <tr className="bg-gray-900 border-b border-gray-800">
                      <th className="p-2 border-r border-gray-800 min-w-[40px]" rowSpan={2}></th>
                      <th colSpan={5} className="p-2 border-r border-gray-800 text-gray-500 font-normal"></th>
                      <th className="p-2 border-r border-gray-800 text-yellow-400 font-medium">I平</th>
                      <th className="p-2 border-r border-gray-800 text-purple-400 font-medium">商务</th>
                      <th className="p-2 border-r border-gray-800 text-gray-400 font-medium">安全</th>
                      <th colSpan={6} className="p-2 border-r border-gray-800 text-orange-400 font-medium text-center">能评及资质</th>
                      <th colSpan={5} className="p-2 border-r border-gray-800 text-cyan-400 font-medium text-center">供应及交付</th>
                      <th colSpan={2} className="p-2 border-r border-gray-800 text-green-400 font-medium text-center">技术</th>
                      <th colSpan={2} className="p-2 border-r border-gray-800 text-pink-400 font-medium text-center">运营能力</th>
                      <th className="p-2 text-red-400 font-medium text-center">问题总结</th>
                    </tr>
                    <tr className="bg-gray-900/50 border-b border-gray-700">
                      {SUMMARY_HEADERS.filter(h => h.key !== 'id').map(h => (
                        <th
                          key={h.key}
                          className="p-2 border-r border-gray-800 text-gray-400 font-medium text-left whitespace-pre-line"
                          style={{ minWidth: h.width * 0.6 }}
                        >
                          {h.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {suppliers.map((supplier, rowIdx) => (
                      <tr key={rowIdx} className="border-b border-gray-800 hover:bg-gray-900/30">
                        <td className="p-2 border-r border-gray-800 text-center">
                          <button
                            onClick={() => deleteRow(rowIdx)}
                            className="text-red-500 hover:text-red-400 text-xs"
                            title="删除此行"
                          >
                            x
                          </button>
                          <div className="text-gray-600 mt-1">{rowIdx + 1}</div>
                        </td>
                        {SUMMARY_HEADERS.filter(h => h.key !== 'id').map(h => (
                          <td
                            key={h.key}
                            className="p-1 border-r border-gray-800 align-top"
                            style={{ minWidth: h.width * 0.6, maxWidth: h.width }}
                          >
                            {editingCell?.row === rowIdx && editingCell?.key === h.key ? (
                              <textarea
                                autoFocus
                                className="w-full bg-gray-800 text-gray-100 p-1 rounded text-xs border border-blue-500 focus:outline-none resize-y"
                                rows={3}
                                value={(supplier as unknown as Record<string, string>)[h.key] || ''}
                                onChange={(e) => handleCellEdit(rowIdx, h.key, e.target.value)}
                                onBlur={() => setEditingCell(null)}
                                onKeyDown={(e) => {
                                  if (e.key === 'Escape') setEditingCell(null);
                                }}
                              />
                            ) : (
                              <div
                                onClick={() => setEditingCell({ row: rowIdx, key: h.key })}
                                className="cursor-text p-1 min-h-[24px] hover:bg-gray-800/50 rounded whitespace-pre-wrap break-words"
                                title="点击编辑"
                              >
                                {(supplier as unknown as Record<string, string>)[h.key] || (
                                  <span className="text-gray-700">-</span>
                                )}
                              </div>
                            )}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* Classification Tab */}
        {activeTab === 'classification' && (
          <div className="max-w-5xl mx-auto space-y-6">
            <h2 className="text-lg font-semibold">腾讯大型数据中心选址分级标准</h2>

            <div className="overflow-x-auto border border-gray-800 rounded-xl">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-900 border-b border-gray-700">
                    <th className="p-3 text-left w-20">级别</th>
                    <th className="p-3 text-left">标准描述（非现货机房、定制化租用）</th>
                    <th className="p-3 text-left">标准描述（现货机房）</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    {
                      level: '预A',
                      color: 'text-green-400 bg-green-500/10',
                      standard: '1、资源能力：能评及合规文件完整；针对全新园区（>100MW）场景，需在东数西算规定节点。\n2、技术能力：技术标准全部满足《技术综合评估表》；\n3、运营能力：建成并运营机架数量50000架（容量350MW）以上或属于现有合作同园区扩容；\n4、交付能力：供应及交付保障在业务指定日期内完成；',
                      spot: '',
                    },
                    {
                      level: '预B',
                      color: 'text-blue-400 bg-blue-500/10',
                      standard: '1、资源能力：能评及合规文件完整；\n2、技术能力：技术标准全部满足《技术综合评估表》；\n3、运营能力：建成并运营机架数量10000架（容量70MW）~50000架（容量350MW）或属于现有合作同园区扩容；\n4、交付能力：供应及交付保障有延期30天内风险；',
                      spot: '1、资源能力：能评及合规文件完整；\n2、技术能力：技术标准部分满足《技术综合评估表》& 不低于GB540174标准A级要求；\n3、运营能力：10000~50000架（70~350MW）；\n4、交付延期30天内风险；\n5、现货品类：非关键设备品类不在白名单；',
                    },
                    {
                      level: '预C',
                      color: 'text-yellow-400 bg-yellow-500/10',
                      standard: '1、资源能力：能评及合规文件完整；不在东数西算规定节点。\n2、技术能力：全部满足《技术综合评估表》；\n3、运营能力：建成并运营机架数量1000~10000架（容量7~70MW）；\n4、交付能力：供应及交付保障有延期60天内风险；',
                      spot: '1、资源能力：能评及合规文件完整；\n2、技术能力：部分满足 & 不低于GB540174 A级；\n3、运营能力：1000~10000架（7~70MW）；\n4、交付延期60天内风险；\n5、关键设备不在白名单须整改；',
                    },
                    {
                      level: '预D',
                      color: 'text-red-400 bg-red-500/10',
                      standard: '1、资源能力：无能评及合规文件缺失；\n2、技术能力：无法满足《技术综合评估表》和GB540174 A级要求；\n3、运营能力：建成并运营机架数量小于1000架（容量7MW）；\n4、交付能力：供应及交付保障有延期60天以上风险；',
                      spot: '',
                    },
                  ].map((row) => (
                    <tr key={row.level} className="border-b border-gray-800 hover:bg-gray-900/30">
                      <td className={`p-3 font-bold ${row.color} rounded-l`}>{row.level}</td>
                      <td className="p-3 whitespace-pre-line text-gray-300">{row.standard}</td>
                      <td className="p-3 whitespace-pre-line text-gray-400">{row.spot || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Key Equipment List */}
            <h3 className="text-md font-semibold mt-8">现货场景关键品类清单</h3>
            <div className="overflow-x-auto border border-gray-800 rounded-xl">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-900 border-b border-gray-700">
                    <th className="p-3 text-left w-16">序号</th>
                    <th className="p-3 text-left">关键设备品类</th>
                    <th className="p-3 text-left">品牌</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    '柴油发电机', '断路器及空气开关', '蓄电池', 'UPS', 'HVDC',
                    '冷水机组', '精密空调末端', '水泵', '电机', '冷却塔',
                  ].map((item, idx) => (
                    <tr key={idx} className="border-b border-gray-800">
                      <td className="p-3 text-gray-500">{idx + 1}</td>
                      <td className="p-3 text-gray-300">{item}</td>
                      <td className="p-3 text-gray-500 text-xs">
                        {idx === 0 ? '以《腾讯合建数据中心主要器件设备推荐选型名单》最新版本品牌为准' : '同上'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
