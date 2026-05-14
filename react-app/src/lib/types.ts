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
  [yr: string]: SalesYearData | string
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

// ── OJC 품번 규칙 ────────────────────────────────────────────
export interface OjcCodeEntry { code: string; label: string }
export interface OjcCountEntry { code: string; count: number }
export interface OjcLengthEntry { m: number; code: string }

export interface OjcRules {
  ojcType:      OjcCodeEntry[]
  coreType:     OjcCodeEntry[]
  coreCount:    OjcCountEntry[]
  material:     OjcCodeEntry[]
  connector:    OjcCodeEntry[]
  marking:      OjcCodeEntry[]
  lengthPreset: OjcLengthEntry[]
}

export const DEFAULT_OJC_RULES: OjcRules = {
  ojcType: [
    { code: 'A', label: 'SOJC / Simplex 0.9mm' },
    { code: 'B', label: 'Simplex 3.0mm' },
    { code: 'C', label: 'DOJC / Duplex 2심' },
    { code: 'D', label: 'MOJC 0.9mm 다심' },
    { code: 'E', label: 'MOJC 2.0mm 다심' },
    { code: 'F', label: 'Pigtail 0.9mm' },
    { code: 'G', label: 'Pigtail 2.0mm' },
    { code: 'H', label: 'Pigtail 3.0mm' },
    { code: 'I', label: 'Pigtail 성단용' },
    { code: 'J', label: '리본케이블' },
    { code: 'K', label: 'DROP-CABLE 3.0mm' },
    { code: 'L', label: 'DROP-CABLE 5.0mm' },
    { code: 'M', label: 'Optical Cable 3.0mm 인입광' },
    { code: 'N', label: 'FLAT 3mm×2mm' },
  ],
  coreType: [
    { code: 'A', label: 'SMF 652D (표준 SM)' },
    { code: 'B', label: 'SMF 657.A1 (밴딩내성 A1)' },
    { code: 'C', label: 'SMF 657.A2 (밴딩내성 A2)' },
    { code: 'D', label: 'SMF 657.B3 (밴딩내성 B3)' },
    { code: 'E', label: 'OM1 (멀티모드 62.5/125)' },
    { code: 'F', label: 'OM2 (멀티모드 50/125)' },
    { code: 'G', label: 'OM3 (멀티모드 50/125 레이저)' },
    { code: 'H', label: 'OM4 (멀티모드 50/125 고급)' },
    { code: 'I', label: 'OM5 (멀티모드 광대역)' },
  ],
  coreCount: [
    { code: '1', count: 1  }, { code: '2', count: 2  },
    { code: '3', count: 4  }, { code: '4', count: 6  },
    { code: '5', count: 7  }, { code: '6', count: 8  },
    { code: '7', count: 10 }, { code: '8', count: 12 },
    { code: '9', count: 13 }, { code: 'A', count: 14 },
    { code: 'B', count: 15 }, { code: 'C', count: 16 },
    { code: 'D', count: 17 }, { code: 'E', count: 18 },
    { code: 'F', count: 20 }, { code: 'G', count: 22 },
    { code: 'H', count: 24 }, { code: 'I', count: 36 },
    { code: 'J', count: 48 },
  ],
  material: [
    { code: '0', label: 'PVC YELLOW'  }, { code: '1', label: 'PVC BLUE'    },
    { code: '2', label: 'PVC AQUA'    }, { code: '3', label: 'PVC ORANGE'  },
    { code: '4', label: 'PVC BLACK'   }, { code: '5', label: 'PVC WHITE'   },
    { code: '6', label: 'PVC 2색'     }, { code: '7', label: 'PVC 4색'     },
    { code: '8', label: 'PVC 6색'     }, { code: '9', label: 'PVC 8색'     },
    { code: 'A', label: 'PVC 12색'    },
    { code: 'B', label: 'LSZH BLACK'  }, { code: 'C', label: 'LSZH YELLOW' },
    { code: 'D', label: 'PU YELLOW'   }, { code: 'E', label: 'PU BLUE'     },
    { code: 'F', label: 'PU AQUA'     }, { code: 'G', label: 'PU ORANGE'   },
    { code: 'H', label: 'PU BLACK'    }, { code: 'I', label: 'PU WHITE'    },
    { code: 'J', label: 'PU 2색'      }, { code: 'K', label: 'PU 4색'      },
    { code: 'L', label: 'PU 6색'      }, { code: 'M', label: 'PU 8색'      },
    { code: 'N', label: 'PU 12색'     },
    { code: 'O', label: 'PE YELLOW'   }, { code: 'P', label: 'PE BLUE'     },
    { code: 'Q', label: 'PE AQUA'     }, { code: 'R', label: 'PE ORANGE'   },
    { code: 'S', label: 'PE BLACK'    }, { code: 'T', label: 'PE WHITE'    },
    { code: 'U', label: 'PE 2색'      }, { code: 'V', label: 'PE 4색'      },
    { code: 'W', label: 'PE 6색'      }, { code: 'X', label: 'PE 8색'      },
    { code: 'Y', label: 'PE 12색'     },
  ],
  connector: [
    { code: '0', label: 'SC/PC'       }, { code: '1', label: 'SC/APC'      },
    { code: '2', label: 'SC/PC CLIP'  }, { code: '3', label: 'SC/APC CLIP' },
    { code: '4', label: 'LC/PC'       }, { code: '5', label: 'LC/APC'      },
    { code: '6', label: 'LC/PC CLIP'  }, { code: '7', label: 'LC/APC CLIP' },
    { code: '8', label: 'FC/PC'       }, { code: '9', label: 'FC/APC'      },
    { code: 'A', label: 'ST/PC'       }, { code: 'B', label: 'ST/APC'      },
    { code: 'C', label: 'MU/PC'       }, { code: 'D', label: 'MU/APC'      },
    { code: 'E', label: 'MTRJ/PC'     }, { code: 'F', label: 'MPO A'       },
    { code: 'G', label: 'MPO B'       }, { code: 'H', label: 'MPO C'       },
    { code: 'I', label: 'MTP A'       }, { code: 'J', label: 'MTP B'       },
    { code: 'K', label: 'MTP C'       },
  ],
  marking: [
    { code: 'N', label: 'NON (마킹없음)' },
    { code: 'K', label: 'KT'            },
    { code: 'L', label: 'LG'            },
    { code: 'O', label: '고려오트론'    },
    { code: 'P', label: '포앤티'        },
    { code: 'Z', label: '수출'          },
    { code: 'X', label: '케이블 사급'   },
  ],
  lengthPreset: [
    { m: 0.5, code: '00E' }, { m: 1,   code: '001' }, { m: 1.5, code: '01E' },
    { m: 2,   code: '002' }, { m: 2.5, code: '02E' }, { m: 3,   code: '003' },
    { m: 4,   code: '004' }, { m: 5,   code: '005' }, { m: 6,   code: '006' },
    { m: 7,   code: '007' }, { m: 8,   code: '008' }, { m: 9,   code: '009' },
    { m: 10,  code: '10J' }, { m: 11,  code: '11J' }, { m: 12,  code: '12J' },
    { m: 13,  code: '13J' }, { m: 14,  code: '14J' }, { m: 15,  code: '15J' },
    { m: 20,  code: '20J' }, { m: 25,  code: '25J' }, { m: 30,  code: '30J' },
    { m: 40,  code: '40J' }, { m: 50,  code: '50J' }, { m: 60,  code: '60J' },
    { m: 75,  code: '75J' }, { m: 100, code: '100' }, { m: 150, code: '150' },
    { m: 200, code: '200' }, { m: 250, code: '250' }, { m: 300, code: '300' },
  ],
}
