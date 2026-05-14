import * as fs from 'fs'
import * as path from 'path'
import { fileURLToPath } from 'url'
import * as XLSX from 'xlsx'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const buf = fs.readFileSync(path.join(__dirname, '..', '검증_가공파일.xlsx'))
const wb  = XLSX.read(buf, { type: 'buffer' })

for (const yr of ['23','24','25','26']) {
  const cSheet = `${yr}년_케이블_집계`
  const hSheet = `${yr}년_하우징_집계`
  if (!wb.SheetNames.includes(cSheet)) continue

  console.log(`\n===== ${cSheet} =====`)
  const cRows = XLSX.utils.sheet_to_json<(string|number)[]>(wb.Sheets[cSheet], { header: 1, defval: 0 }) as (string|number)[][]
  console.log('케이블 타입            파이      단위   연간 합계   월 최대  최대월')
  for (const r of cRows.slice(1)) {
    if (!r[0]) continue
    console.log(
      String(r[0]).padEnd(22) +
      String(r[1]).padEnd(10) +
      String(r[2]).padEnd(6) +
      String(r[15]).padStart(10) +
      String(r[16]).padStart(9) +
      '  ' + String(r[17])
    )
  }

  if (!wb.SheetNames.includes(hSheet)) continue
  console.log(`\n===== ${hSheet} =====`)
  const hRows = XLSX.utils.sheet_to_json<(string|number)[]>(wb.Sheets[hSheet], { header: 1, defval: 0 }) as (string|number)[][]
  console.log('하우징 타입                      연간 합계  월 최대  최대월')
  for (const r of hRows.slice(1)) {
    if (!r[0]) continue
    console.log(
      String(r[0]).padEnd(33) +
      String(r[13]).padStart(9) +
      String(r[14]).padStart(8) +
      '  ' + String(r[15])
    )
  }
}
