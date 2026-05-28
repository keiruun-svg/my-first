import { useState, useEffect, useMemo } from 'react'
import {
  loadImportOrders, saveImportOrders, loadVendors,
} from '../lib/supabase'
import type { ImportOrder, Vendor } from '../lib/supabase'

function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2)
}

function orderLabel(o: ImportOrder) {
  return `${o.vendorCode} ${o.year}-${String(o.seq).padStart(2, '0')}차`
}

function dday(expectedDate: string): number {
  const today = new Date(); today.setHours(0, 0, 0, 0)
  const exp   = new Date(expectedDate)
  return Math.round((exp.getTime() - today.getTime()) / 86400000)
}

function leadDays(orderDate: string, actualDate: string): number {
  return Math.round(
    (new Date(actualDate).getTime() - new Date(orderDate).getTime()) / 86400000
  )
}

function todayStr() {
  return new Date().toISOString().slice(0, 10)
}

export default function ImportOrdersView({ prefill }: { prefill?: { note: string } }) {
  const [orders,   setOrders]   = useState<ImportOrder[]>([])
  const [vendors,  setVendors]  = useState<Vendor[]>([])
  const [filter,   setFilter]   = useState<'all' | 'pending' | 'done'>('all')
  const [editId,   setEditId]   = useState<string | null>(null)  // 입고일 인라인 입력 중인 ID
  const [editDate, setEditDate] = useState('')

  // 등록 폼 상태
  const currentYear = String(new Date().getFullYear()).slice(2)
  const [form, setForm] = useState({
    vendorCode:   '',
    year:         currentYear,
    seq:          1,
    orderDate:    todayStr(),
    expectedDate: '',
    note:         '',
  })

  useEffect(() => {
    if (prefill) setForm(f => ({ ...f, note: prefill.note }))
  }, [prefill])

  useEffect(() => {
    setOrders(loadImportOrders())
    loadVendors().then(setVendors)
  }, [])

  // 선택 업체의 다음 차수 자동 제안
  useEffect(() => {
    if (!form.vendorCode || !form.year) return
    const maxSeq = orders
      .filter(o => o.vendorCode === form.vendorCode && o.year === form.year)
      .reduce((m, o) => Math.max(m, o.seq), 0)
    setForm(f => ({ ...f, seq: maxSeq + 1 }))
  }, [form.vendorCode, form.year, orders])

  function handleRegister() {
    if (!form.vendorCode || !form.orderDate || !form.expectedDate) return
    const vendor = vendors.find(v => v.code === form.vendorCode)
    const newOrder: ImportOrder = {
      id:           genId(),
      vendorCode:   form.vendorCode,
      vendorName:   vendor?.name ?? form.vendorCode,
      year:         form.year,
      seq:          form.seq,
      orderDate:    form.orderDate,
      expectedDate: form.expectedDate,
      note:         form.note || undefined,
    }
    const updated = [...orders, newOrder].sort((a, b) =>
      b.orderDate.localeCompare(a.orderDate)
    )
    setOrders(updated)
    saveImportOrders(updated)
    // 등록 후 차수 +1, 날짜만 유지
    setForm(f => ({ ...f, seq: f.seq + 1, expectedDate: '', note: '' }))
  }

  function handleReceive(id: string) {
    if (!editDate) return
    const updated = orders.map(o =>
      o.id === id ? { ...o, actualDate: editDate } : o
    )
    setOrders(updated)
    saveImportOrders(updated)
    setEditId(null)
    setEditDate('')
  }

  function handleDelete(id: string) {
    if (!confirm('이 발주건을 삭제할까요?')) return
    const updated = orders.filter(o => o.id !== id)
    setOrders(updated)
    saveImportOrders(updated)
  }

  const displayed = useMemo(() => {
    return orders.filter(o => {
      if (filter === 'pending') return !o.actualDate
      if (filter === 'done')    return  !!o.actualDate
      return true
    })
  }, [orders, filter])

  const pendingCount = orders.filter(o => !o.actualDate).length
  const doneCount    = orders.filter(o =>  !!o.actualDate).length

  const filterBtn = (id: typeof filter, label: string) => (
    <button
      onClick={() => setFilter(id)}
      className={`px-3 py-1.5 text-xs rounded-full font-medium transition ${
        filter === id
          ? 'bg-blue-600 text-white'
          : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
      }`}
    >
      {label}
    </button>
  )

  return (
    <div className="space-y-5">
      {/* 등록 폼 */}
      <div className="bg-gray-50 border rounded-lg p-4 space-y-3">
        <h3 className="text-sm font-semibold text-gray-700">발주 등록</h3>
        <div className="flex flex-wrap gap-3 items-end">
          {/* 업체 */}
          <div className="space-y-1">
            <label className="text-xs text-gray-500">업체</label>
            <select
              value={form.vendorCode}
              onChange={e => setForm(f => ({ ...f, vendorCode: e.target.value }))}
              className="border rounded px-2 py-1.5 text-sm min-w-28"
            >
              <option value="">선택</option>
              {vendors.map(v => (
                <option key={v.code} value={v.code}>{v.code} — {v.name}</option>
              ))}
            </select>
          </div>

          {/* 연도 */}
          <div className="space-y-1">
            <label className="text-xs text-gray-500">연도(YY)</label>
            <input
              type="text" maxLength={2} value={form.year}
              onChange={e => setForm(f => ({ ...f, year: e.target.value }))}
              className="border rounded px-2 py-1.5 text-sm w-16 text-center"
              placeholder="26"
            />
          </div>

          {/* 차수 */}
          <div className="space-y-1">
            <label className="text-xs text-gray-500">차수</label>
            <input
              type="number" min={1} max={99} value={form.seq}
              onChange={e => setForm(f => ({ ...f, seq: parseInt(e.target.value) || 1 }))}
              className="border rounded px-2 py-1.5 text-sm w-16 text-center"
            />
          </div>

          {/* 발주일 */}
          <div className="space-y-1">
            <label className="text-xs text-gray-500">발주일</label>
            <input
              type="date" value={form.orderDate}
              onChange={e => setForm(f => ({ ...f, orderDate: e.target.value }))}
              className="border rounded px-2 py-1.5 text-sm"
            />
          </div>

          {/* 예상입고일 */}
          <div className="space-y-1">
            <label className="text-xs text-gray-500">예상 입고일</label>
            <input
              type="date" value={form.expectedDate}
              onChange={e => setForm(f => ({ ...f, expectedDate: e.target.value }))}
              className="border rounded px-2 py-1.5 text-sm"
            />
          </div>

          <button
            onClick={handleRegister}
            disabled={!form.vendorCode || !form.orderDate || !form.expectedDate}
            className="px-4 py-1.5 bg-blue-600 text-white text-sm font-medium rounded hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed transition"
          >
            등록
          </button>
        </div>

        {/* 메모 */}
        <div className="space-y-1">
          <label className="text-xs text-gray-500">메모 (발주 품목 등)</label>
          <input
            type="text" value={form.note}
            onChange={e => setForm(f => ({ ...f, note: e.target.value }))}
            className="border rounded px-2 py-1.5 text-sm w-full max-w-xl"
            placeholder="예: 14-K-107 × 200개, 14-P-201 × 100개"
          />
        </div>

        {/* 미리보기 */}
        {form.vendorCode && form.year && form.seq && (
          <p className="text-xs text-gray-400">
            발주건명: <span className="font-medium text-gray-600">
              {form.vendorCode} {form.year}-{String(form.seq).padStart(2, '0')}차
            </span>
          </p>
        )}
      </div>

      {/* 필터 + 요약 */}
      <div className="flex items-center gap-2 flex-wrap">
        {filterBtn('all',     `전체 ${orders.length}`)}
        {filterBtn('pending', `⏳ 입고대기 ${pendingCount}`)}
        {filterBtn('done',    `✅ 완료 ${doneCount}`)}
      </div>

      {/* 현황 테이블 */}
      {displayed.length === 0 ? (
        <div className="text-center py-10 text-gray-400 text-sm">
          {orders.length === 0 ? '등록된 발주건이 없습니다.' : '해당하는 발주건이 없습니다.'}
        </div>
      ) : (
        <div className="overflow-x-auto rounded-lg border">
          <table className="min-w-full text-sm">
            <thead>
              <tr className="bg-[#1F3864] text-white text-xs">
                <th className="px-3 py-2 text-left whitespace-nowrap">발주건</th>
                <th className="px-3 py-2 text-center whitespace-nowrap">발주일</th>
                <th className="px-3 py-2 text-center whitespace-nowrap">예상 입고일</th>
                <th className="px-3 py-2 text-center whitespace-nowrap">D-day</th>
                <th className="px-3 py-2 text-center whitespace-nowrap">실제 입고일</th>
                <th className="px-3 py-2 text-center whitespace-nowrap">리드타임</th>
                <th className="px-3 py-2 text-center whitespace-nowrap">상태</th>
                <th className="px-3 py-2 text-center whitespace-nowrap"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {displayed.map((order, i) => {
                const done = !!order.actualDate
                const dd   = !done ? dday(order.expectedDate) : null
                const lt   = done ? leadDays(order.orderDate, order.actualDate!) : null
                const overdue = dd !== null && dd < 0

                return (
                  <tr key={order.id} className={i % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    {/* 발주건명 */}
                    <td className="px-3 py-2">
                      <div className="font-medium whitespace-nowrap">
                        {orderLabel(order)}
                        <span className="ml-2 text-xs text-gray-400">{order.vendorName}</span>
                      </div>
                      {order.note && (
                        <div className="text-xs text-blue-600 mt-0.5 max-w-xs truncate" title={order.note}>
                          {order.note}
                        </div>
                      )}
                    </td>

                    {/* 발주일 */}
                    <td className="px-3 py-2 text-center text-gray-600 whitespace-nowrap">
                      {order.orderDate}
                    </td>

                    {/* 예상 입고일 */}
                    <td className="px-3 py-2 text-center text-gray-600 whitespace-nowrap">
                      {order.expectedDate}
                    </td>

                    {/* D-day */}
                    <td className="px-3 py-2 text-center whitespace-nowrap">
                      {done ? (
                        <span className="text-gray-300 text-xs">—</span>
                      ) : dd === 0 ? (
                        <span className="font-bold text-blue-600">D-day</span>
                      ) : dd !== null && dd > 0 ? (
                        <span className="text-gray-700">D-{dd}</span>
                      ) : (
                        <span className="font-medium text-red-600">D+{Math.abs(dd!)}</span>
                      )}
                    </td>

                    {/* 실제 입고일 */}
                    <td className="px-3 py-2 text-center whitespace-nowrap">
                      {done ? (
                        <span className="text-gray-700">{order.actualDate}</span>
                      ) : editId === order.id ? (
                        <div className="flex items-center gap-1 justify-center">
                          <input
                            type="date"
                            value={editDate}
                            onChange={e => setEditDate(e.target.value)}
                            className="border rounded px-1.5 py-0.5 text-xs w-32"
                            autoFocus
                          />
                          <button
                            onClick={() => handleReceive(order.id)}
                            disabled={!editDate}
                            className="text-xs px-2 py-0.5 bg-green-600 text-white rounded disabled:opacity-40"
                          >
                            확인
                          </button>
                          <button
                            onClick={() => { setEditId(null); setEditDate('') }}
                            className="text-xs px-1.5 py-0.5 text-gray-400 hover:text-gray-600"
                          >
                            ✕
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => { setEditId(order.id); setEditDate(todayStr()) }}
                          className={`text-xs px-2 py-0.5 rounded border transition ${
                            overdue
                              ? 'border-red-300 text-red-600 hover:bg-red-50'
                              : 'border-gray-300 text-gray-500 hover:bg-gray-100'
                          }`}
                        >
                          입고 처리
                        </button>
                      )}
                    </td>

                    {/* 리드타임 */}
                    <td className="px-3 py-2 text-center whitespace-nowrap">
                      {lt !== null ? (
                        <span className="text-gray-700">{lt}일</span>
                      ) : (
                        <span className="text-gray-300 text-xs">—</span>
                      )}
                    </td>

                    {/* 상태 */}
                    <td className="px-3 py-2 text-center whitespace-nowrap">
                      {done ? (
                        <span className="inline-block text-xs font-medium px-2 py-0.5 rounded-full text-green-700 bg-green-100">
                          ✅ 완료
                        </span>
                      ) : overdue ? (
                        <span className="inline-block text-xs font-medium px-2 py-0.5 rounded-full text-red-700 bg-red-100">
                          지연
                        </span>
                      ) : (
                        <span className="inline-block text-xs font-medium px-2 py-0.5 rounded-full text-yellow-700 bg-yellow-100">
                          ⏳ 대기
                        </span>
                      )}
                    </td>

                    {/* 삭제 */}
                    <td className="px-3 py-2 text-center">
                      <button
                        onClick={() => handleDelete(order.id)}
                        className="text-gray-300 hover:text-red-400 text-xs transition"
                        title="삭제"
                      >
                        ✕
                      </button>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
