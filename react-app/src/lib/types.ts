export type YearKey = '23' | '24' | '25'

export interface YearStats {
  monthly: number[]
  annual: number
  peak: number
}

export interface CableMeta {
  품번: string
  품명: string
  구매처: string
  리드타임: string | number | null
}

export interface HousingComp {
  품번: string
  품명: string
  구매처: string
  리드타임: string | number | null
}

export interface InventoryCableItem {
  현재고: number
}

export interface InventoryHousingItem {
  현재고: number
  기발주: number
}

export interface Metadata {
  cable: Record<string, CableMeta>
  housing: Record<string, HousingComp | HousingComp[]>
  ferrule?: Record<string, { 품번: string; 품명: string; 구매처: string; 리드타임: string | null }>
}

export interface Inventory {
  cable: Record<string, InventoryCableItem>
  housing: Record<string, InventoryHousingItem | InventoryHousingItem[]>
}

export interface SalesYearData {
  sales: number
  production: number
  ratio: number
}

export interface SalesItem {
  품목명: string
  '23': SalesYearData
  '24': SalesYearData
  '25': SalesYearData
}

export type SalesAnalysis = Record<string, SalesItem>

export interface AppSettings {
  lead_time_default: number
  colors: {
    main_header: string
    year_23: string
    year_24: string
    year_25: string
  }
}

export const DEFAULT_SETTINGS: AppSettings = {
  lead_time_default: 60,
  colors: {
    main_header: '1F3864',
    year_23: '2F5597',
    year_24: '2E75B6',
    year_25: '155480',
  },
}

export type CableKey = [string, string]
export type HousingKey = [string, string]
