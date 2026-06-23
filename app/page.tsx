'use client'

import { useState, useRef, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { getSupabase, type Winner } from '../lib/supabase'

type Tab = 'survey' | 'lucky' | 'results'
type Mode = 'setup' | 'draw'
type LuckyPrize = '논픽션 핸드크림' | '하이드로 텀블러' | 'TWG Tea'

interface Participant {
  key: string
  drumName: string
  display: string
}

const LUCKY_PRIZES: { name: LuckyPrize; count: number }[] = [
  { name: '논픽션 핸드크림', count: 5 },
  { name: '하이드로 텀블러', count: 3 },
  { name: 'TWG Tea', count: 3 },
]

function shuffle<T>(arr: T[]): T[] {
  const a = [...arr]
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[a[i], a[j]] = [a[j], a[i]]
  }
  return a
}

function last4(phone: string): string {
  const digits = phone.replace(/\D/g, '')
  return digits.slice(-4)
}

function parseSurveyExcel(data: Uint8Array): Participant[] {
  const workbook = XLSX.read(data, { type: 'array' })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows: string[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 })
  return rows
    .slice(1)
    .filter((row) => row[2] && String(row[2]).trim())
    .map((row) => {
      const type = String(row[0] ?? '').trim()
      const region = String(row[1] ?? '').trim()
      const idInfo = String(row[2] ?? '').trim()
      return {
        key: idInfo,
        drumName: idInfo,
        display: `${type} | ${region} | ${idInfo}`,
      }
    })
}

function parseLuckyExcel(data: Uint8Array): Participant[] {
  const workbook = XLSX.read(data, { type: 'array' })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows: string[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 })
  return rows
    .slice(1)
    .filter((row) => row[1] && String(row[1]).trim())
    .map((row) => {
      const name = String(row[1] ?? '').trim()
      const phone = String(row[2] ?? '').trim()
      const phoneLast4 = last4(phone)
      return {
        key: `${name}_${phoneLast4}`,
        drumName: name,
        display: phoneLast4 ? `${name} (${phoneLast4})` : name,
      }
    })
}

export default function Home() {
  const [mode, setMode] = useState<Mode>('setup')
  const [activeTab, setActiveTab] = useState<Tab>('survey')

  // Survey state
  const [surveyParticipants, setSurveyParticipants] = useState<Participant[]>([])
  const [surveyFileName, setSurveyFileName] = useState('')
  const [surveyCount, setSurveyCount] = useState(10)
  const [surveyDrawing, setSurveyDrawing] = useState(false)
  const [surveyDrumName, setSurveyDrumName] = useState('')
  const [surveyPending, setSurveyPending] = useState<Participant[]>([])
  const [surveyConfirmed, setSurveyConfirmed] = useState<Participant[]>([])
  const surveyTimerRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Lucky draw state
  const [luckyParticipants, setLuckyParticipants] = useState<Participant[]>([])
  const [luckyFileName, setLuckyFileName] = useState('')
  const [luckyPrizeTab, setLuckyPrizeTab] = useState<LuckyPrize>('논픽션 핸드크림')
  const [luckyDrawing, setLuckyDrawing] = useState(false)
  const [luckyDrumName, setLuckyDrumName] = useState('')
  const [luckyPending, setLuckyPending] = useState<Participant[]>([])
  const [luckyConfirmed, setLuckyConfirmed] = useState<Record<LuckyPrize, Participant[]>>({
    '논픽션 핸드크림': [],
    '하이드로 텀블러': [],
    'TWG Tea': [],
  })
  const luckyTimerRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Results state
  const [allWinners, setAllWinners] = useState<Winner[]>([])
  const [copied, setCopied] = useState(false)

  const getAllLuckyConfirmedKeys = useCallback(() => {
    return new Set(Object.values(luckyConfirmed).flat().map((p) => p.key))
  }, [luckyConfirmed])

  const getSurveyPool = useCallback(() => {
    const excludedKeys = new Set(surveyConfirmed.map((p) => p.key))
    return surveyParticipants.filter((p) => !excludedKeys.has(p.key))
  }, [surveyParticipants, surveyConfirmed])

  const getLuckyPool = useCallback(() => {
    const excludedKeys = getAllLuckyConfirmedKeys()
    return luckyParticipants.filter((p) => !excludedKeys.has(p.key))
  }, [luckyParticipants, getAllLuckyConfirmedKeys])

  const currentPrizeConfig = LUCKY_PRIZES.find((p) => p.name === luckyPrizeTab)!
  const currentPrizeRemaining = currentPrizeConfig.count - luckyConfirmed[luckyPrizeTab].length

  const startDrumRoll = (
    pool: Participant[],
    count: number,
    setDrumNameFn: (n: string) => void,
    setDrawing: (b: boolean) => void,
    setPending: (p: Participant[]) => void,
    timerRef: React.MutableRefObject<ReturnType<typeof setInterval> | null>
  ) => {
    if (pool.length === 0) return
    const drawCount = Math.min(count, pool.length)
    setDrawing(true)
    setPending([])
    setDrumNameFn(pool[Math.floor(Math.random() * pool.length)].drumName)

    let tick = 0
    const totalTicks = 30

    const run = () => {
      tick++
      setDrumNameFn(pool[Math.floor(Math.random() * pool.length)].drumName)
      if (tick >= totalTicks) {
        if (timerRef.current) clearInterval(timerRef.current)
        const winners = shuffle(pool).slice(0, drawCount)
        setPending(winners)
        setDrumNameFn('')
        setDrawing(false)
      }
    }

    timerRef.current = setInterval(run, 100)
  }

  // Survey handlers
  const handleSurveyDraw = () => {
    const pool = getSurveyPool()
    setSurveyPending([])
    startDrumRoll(pool, surveyCount, setSurveyDrumName, setSurveyDrawing, setSurveyPending, surveyTimerRef)
  }

  const handleSurveyConfirm = async () => {
    setSurveyConfirmed((prev) => [...prev, ...surveyPending])
    const rows = surveyPending.map((p) => ({
      name: p.display,
      prize: '배달의민족 상품권 5만원권',
      prize_type: 'survey' as const,
    }))
    await getSupabase().from('winners').insert(rows)
    setSurveyPending([])
  }

  const handleSurveyRedraw = () => {
    const pendingKeys = new Set(surveyPending.map((p) => p.key))
    const pool = getSurveyPool().filter((p) => !pendingKeys.has(p.key))
    setSurveyPending([])
    startDrumRoll(pool, surveyCount, setSurveyDrumName, setSurveyDrawing, setSurveyPending, surveyTimerRef)
  }

  const handleSurveyRedrawOne = (idx: number) => {
    const excludedKeys = new Set([
      ...surveyConfirmed.map((p) => p.key),
      ...surveyPending.map((p) => p.key),
    ])
    const pool = surveyParticipants.filter((p) => !excludedKeys.has(p.key))
    if (pool.length === 0) return
    const replacement = shuffle(pool)[0]
    setSurveyPending((prev) => prev.map((p, i) => (i === idx ? replacement : p)))
  }

  // Lucky draw handlers
  const handleLuckyDraw = () => {
    const pool = getLuckyPool()
    setLuckyPending([])
    startDrumRoll(pool, currentPrizeRemaining, setLuckyDrumName, setLuckyDrawing, setLuckyPending, luckyTimerRef)
  }

  const handleLuckyConfirm = async () => {
    setLuckyConfirmed((prev) => ({
      ...prev,
      [luckyPrizeTab]: [...prev[luckyPrizeTab], ...luckyPending],
    }))
    const rows = luckyPending.map((p) => ({
      name: p.display,
      prize: luckyPrizeTab,
      prize_type: 'lucky' as const,
    }))
    await getSupabase().from('winners').insert(rows)
    setLuckyPending([])
  }

  const handleLuckyRedrawOne = (idx: number) => {
    const allConfirmedKeys = getAllLuckyConfirmedKeys()
    const pendingKeys = new Set(luckyPending.map((p) => p.key))
    const pool = luckyParticipants.filter(
      (p) => !allConfirmedKeys.has(p.key) && !pendingKeys.has(p.key)
    )
    if (pool.length === 0) return
    const replacement = shuffle(pool)[0]
    setLuckyPending((prev) => prev.map((p, i) => (i === idx ? replacement : p)))
  }

  // File upload handlers
  const handleSurveyFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setSurveyFileName(file.name)
    const reader = new FileReader()
    reader.onload = (ev) => {
      const data = new Uint8Array(ev.target?.result as ArrayBuffer)
      setSurveyParticipants(parseSurveyExcel(data))
    }
    reader.readAsArrayBuffer(file)
  }

  const handleLuckyFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setLuckyFileName(file.name)
    const reader = new FileReader()
    reader.onload = (ev) => {
      const data = new Uint8Array(ev.target?.result as ArrayBuffer)
      setLuckyParticipants(parseLuckyExcel(data))
    }
    reader.readAsArrayBuffer(file)
  }

  // Results
  const loadResults = async () => {
    const { data } = await getSupabase()
      .from('winners')
      .select('*')
      .order('confirmed_at', { ascending: true })
    if (data) setAllWinners(data)
  }

  useEffect(() => {
    if (activeTab === 'results') loadResults()
  }, [activeTab])

  const copyResults = () => {
    const surveyWinners = allWinners.filter((w) => w.prize_type === 'survey')
    const luckyWinners = allWinners.filter((w) => w.prize_type === 'lucky')
    let text = '=== 설문조사 경품추첨 당첨자 ===\n'
    surveyWinners.forEach((w, i) => {
      text += `${i + 1}. ${w.name}\n`
    })
    text += '\n=== 럭키드로우 당첨자 ===\n'
    const grouped: Record<string, string[]> = {}
    luckyWinners.forEach((w) => {
      if (!grouped[w.prize]) grouped[w.prize] = []
      grouped[w.prize].push(w.name)
    })
    Object.entries(grouped).forEach(([prize, names]) => {
      text += `\n[${prize}]\n`
      names.forEach((n, i) => {
        text += `${i + 1}. ${n}\n`
      })
    })
    navigator.clipboard.writeText(text)
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }

  const canStartDraw = surveyParticipants.length > 0 || luckyParticipants.length > 0

  // ==================== SETUP MODE ====================
  if (mode === 'setup') {
    return (
      <div className="min-h-screen bg-gray-50">
        <header className="bg-gray-800 text-white py-4 px-4">
          <h1 className="text-center text-xl font-bold">추첨 관리자 설정</h1>
          <p className="text-center text-sm mt-1 opacity-75">
            명단 업로드 후 &quot;추첨 화면으로 전환&quot; 버튼을 누르세요
          </p>
        </header>

        <main className="max-w-2xl mx-auto p-6 space-y-6">
          {/* Survey file upload */}
          <div className="bg-white rounded-xl shadow-md p-6">
            <h2 className="text-lg font-bold text-gray-800 mb-2">1. 설문조사 추첨 명단</h2>
            <p className="text-sm text-gray-500 mb-4">
              엑셀 형식: A열(센터유형), B열(지역), C열(아이디+연락처 뒷번호)
            </p>
            <div className="flex items-center gap-3">
              <label className="text-sm text-white cursor-pointer bg-gray-700 hover:bg-gray-800 px-4 py-2 rounded-lg transition-colors font-medium">
                파일 선택
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={handleSurveyFileUpload}
                />
              </label>
              {surveyFileName && (
                <span className="text-sm text-gray-600">{surveyFileName}</span>
              )}
            </div>
            {surveyParticipants.length > 0 && (
              <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg">
                <span className="text-sm font-medium text-green-700">
                  {surveyParticipants.length}명 로드 완료
                </span>
                <div className="mt-2 max-h-40 overflow-y-auto text-xs text-gray-600 space-y-1">
                  {surveyParticipants.map((p, i) => (
                    <div key={i} className="py-0.5">{i + 1}. {p.display}</div>
                  ))}
                </div>
              </div>
            )}
            <div className="mt-4">
              <label className="text-sm font-medium text-gray-700">추첨 인원</label>
              <div className="flex items-center gap-3 mt-1">
                <input
                  type="number"
                  min={1}
                  max={Math.max(1, surveyParticipants.length)}
                  value={surveyCount}
                  onChange={(e) => setSurveyCount(Number(e.target.value))}
                  className="w-20 border border-gray-300 rounded-lg px-3 py-2 text-sm"
                />
                <span className="text-sm text-gray-500">명</span>
              </div>
            </div>
          </div>

          {/* Lucky draw file upload */}
          <div className="bg-white rounded-xl shadow-md p-6">
            <h2 className="text-lg font-bold text-gray-800 mb-2">2. 럭키드로우 명단</h2>
            <p className="text-sm text-gray-500 mb-4">
              엑셀 형식: B열(이름), C열(연락처) — 당첨 시 이름 + 연락처 뒷4자리 표시
            </p>
            <div className="flex items-center gap-3">
              <label className="text-sm text-white cursor-pointer bg-gray-700 hover:bg-gray-800 px-4 py-2 rounded-lg transition-colors font-medium">
                파일 선택
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  className="hidden"
                  onChange={handleLuckyFileUpload}
                />
              </label>
              {luckyFileName && (
                <span className="text-sm text-gray-600">{luckyFileName}</span>
              )}
            </div>
            {luckyParticipants.length > 0 && (
              <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg">
                <span className="text-sm font-medium text-green-700">
                  {luckyParticipants.length}명 로드 완료
                </span>
                <div className="mt-2 max-h-40 overflow-y-auto text-xs text-gray-600 space-y-1">
                  {luckyParticipants.map((p, i) => (
                    <div key={i} className="py-0.5">{i + 1}. {p.display}</div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Switch to draw mode */}
          <button
            onClick={() => setMode('draw')}
            disabled={!canStartDraw}
            className="w-full bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-lg font-bold py-4 rounded-xl transition-colors shadow-lg"
          >
            {canStartDraw ? '추첨 화면으로 전환' : '명단을 먼저 업로드하세요'}
          </button>
        </main>
      </div>
    )
  }

  // ==================== DRAW MODE ====================
  const drawTabs: { key: Tab; label: string }[] = [
    { key: 'survey', label: '설문조사 추첨' },
    { key: 'lucky', label: '럭키드로우' },
    { key: 'results', label: '최종 결과' },
  ]

  return (
    <div className="min-h-screen bg-gradient-to-b from-[#e8f4fd] to-white">
      <header className="bg-primary text-white py-4 px-4 shadow-lg relative">
        <h1 className="text-center text-xl md:text-2xl font-bold">
          2026년 국가·지역진로교육센터 및 진로체험지원센터 워크숍
        </h1>
        <p className="text-center text-sm md:text-base mt-1 opacity-90">경품추첨 이벤트</p>
        <button
          onClick={() => setMode('setup')}
          className="absolute top-4 right-4 text-white/40 hover:text-white/80 text-xs transition-colors"
          title="관리자 설정"
        >
          설정
        </button>
      </header>

      <nav className="flex border-b border-gray-200 bg-white sticky top-0 z-10">
        {drawTabs.map((tab) => (
          <button
            key={tab.key}
            onClick={() => setActiveTab(tab.key)}
            className={`flex-1 py-3 px-4 text-sm md:text-base font-medium transition-colors ${
              activeTab === tab.key
                ? 'text-primary border-b-2 border-primary bg-primary-light/50'
                : 'text-gray-500 hover:text-gray-700'
            }`}
          >
            {tab.label}
          </button>
        ))}
      </nav>

      <main className="max-w-3xl mx-auto p-4 md:p-6">
        {/* Tab 1: Survey Draw */}
        {activeTab === 'survey' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl shadow-md p-6 text-center min-h-[300px] flex flex-col items-center justify-center">
              {surveyDrawing && (
                <div>
                  <div className="text-4xl md:text-6xl font-bold text-primary drumroll-active py-8">
                    {surveyDrumName || '...'}
                  </div>
                  <p className="text-gray-500 mt-2 text-lg">추첨 중...</p>
                </div>
              )}

              {!surveyDrawing && surveyPending.length === 0 && (
                <div>
                  <p className="text-gray-400 mb-6 text-lg">
                    설문에 참여해주신 모든 분들께 감사드립니다.
                  </p>
                  <p className="text-primary font-bold text-xl mb-8">
                    배달의민족 상품권 5만원권 ({surveyCount}명)
                  </p>
                  <button
                    onClick={handleSurveyDraw}
                    disabled={getSurveyPool().length === 0}
                    className="bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-xl font-bold py-5 px-16 rounded-full transition-colors shadow-lg hover:shadow-xl"
                  >
                    추첨 시작
                  </button>
                </div>
              )}

              {!surveyDrawing && surveyPending.length > 0 && (
                <div className="w-full">
                  <h3 className="text-2xl font-bold text-gray-700 mb-6">당첨을 축하합니다!</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-8">
                    {surveyPending.map((p, i) => (
                      <div
                        key={i}
                        className="winner-reveal bg-gradient-to-br from-primary to-primary-dark text-white rounded-xl p-5 shadow-lg text-left"
                        style={{ animationDelay: `${i * 0.1}s` }}
                      >
                        <div className="text-lg font-bold">{p.display}</div>
                        <button
                          onClick={() => handleSurveyRedrawOne(i)}
                          className="text-xs mt-2 opacity-60 hover:opacity-100 underline"
                        >
                          재추첨
                        </button>
                      </div>
                    ))}
                  </div>
                  <div className="flex gap-3 justify-center">
                    <button
                      onClick={handleSurveyConfirm}
                      className="bg-green-500 hover:bg-green-600 text-white font-bold py-3 px-8 rounded-full transition-colors text-lg"
                    >
                      당첨 확인
                    </button>
                    <button
                      onClick={handleSurveyRedraw}
                      className="bg-gray-400 hover:bg-gray-500 text-white font-bold py-3 px-8 rounded-full transition-colors"
                    >
                      전체 재추첨
                    </button>
                  </div>
                </div>
              )}
            </div>

            {surveyConfirmed.length > 0 && (
              <div className="bg-white rounded-xl shadow-md p-6">
                <h2 className="text-lg font-bold text-primary mb-4">
                  확정 당첨자 ({surveyConfirmed.length}명)
                </h2>
                <div className="space-y-2">
                  {surveyConfirmed.map((p, i) => (
                    <div
                      key={i}
                      className="bg-primary-light text-primary rounded-lg p-3 text-sm font-medium"
                    >
                      {i + 1}. {p.display}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Tab 2: Lucky Draw */}
        {activeTab === 'lucky' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl shadow-md overflow-hidden">
              <div className="flex border-b">
                {LUCKY_PRIZES.map((prize) => (
                  <button
                    key={prize.name}
                    onClick={() => {
                      setLuckyPrizeTab(prize.name)
                      setLuckyPending([])
                    }}
                    className={`flex-1 py-3 px-2 text-sm font-medium transition-colors ${
                      luckyPrizeTab === prize.name
                        ? 'text-primary border-b-2 border-primary bg-primary-light/50'
                        : 'text-gray-500 hover:text-gray-700'
                    }`}
                  >
                    {prize.name}
                    <br />
                    <span className="text-xs">
                      ({luckyConfirmed[prize.name].length}/{prize.count}명)
                    </span>
                  </button>
                ))}
              </div>

              <div className="p-6 text-center min-h-[300px] flex flex-col items-center justify-center">
                {luckyDrawing && (
                  <div>
                    <div className="text-4xl md:text-6xl font-bold text-primary drumroll-active py-8">
                      {luckyDrumName || '...'}
                    </div>
                    <p className="text-gray-500 mt-2 text-lg">추첨 중...</p>
                  </div>
                )}

                {!luckyDrawing && luckyPending.length === 0 && (
                  <div>
                    <p className="text-gray-400 mb-6 text-lg">행운의 주인공을 기다립니다!</p>
                    <p className="text-primary font-bold text-xl mb-8">
                      {luckyPrizeTab} ({currentPrizeRemaining}명 남음)
                    </p>
                    <button
                      onClick={handleLuckyDraw}
                      disabled={currentPrizeRemaining <= 0 || getLuckyPool().length === 0}
                      className="bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-xl font-bold py-5 px-16 rounded-full transition-colors shadow-lg hover:shadow-xl"
                    >
                      {currentPrizeRemaining <= 0 ? '추첨 완료' : '추첨 시작'}
                    </button>
                  </div>
                )}

                {!luckyDrawing && luckyPending.length > 0 && (
                  <div className="w-full">
                    <h3 className="text-2xl font-bold text-gray-700 mb-6">
                      {luckyPrizeTab} 당첨을 축하합니다!
                    </h3>
                    <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-8">
                      {luckyPending.map((p, i) => (
                        <div
                          key={i}
                          className="winner-reveal bg-gradient-to-br from-primary to-primary-dark text-white rounded-xl p-4 shadow-lg"
                          style={{ animationDelay: `${i * 0.1}s` }}
                        >
                          <div className="text-lg font-bold">{p.display}</div>
                          <button
                            onClick={() => handleLuckyRedrawOne(i)}
                            className="text-xs mt-2 opacity-60 hover:opacity-100 underline"
                          >
                            재추첨
                          </button>
                        </div>
                      ))}
                    </div>
                    <div className="flex gap-3 justify-center">
                      <button
                        onClick={handleLuckyConfirm}
                        className="bg-green-500 hover:bg-green-600 text-white font-bold py-3 px-8 rounded-full transition-colors text-lg"
                      >
                        당첨 확인
                      </button>
                      <button
                        onClick={handleLuckyDraw}
                        className="bg-gray-400 hover:bg-gray-500 text-white font-bold py-3 px-8 rounded-full transition-colors"
                      >
                        전체 재추첨
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>

            {Object.values(luckyConfirmed).flat().length > 0 && (
              <div className="bg-white rounded-xl shadow-md p-6">
                <h2 className="text-lg font-bold text-primary mb-4">럭키드로우 당첨자</h2>
                {LUCKY_PRIZES.map(
                  (prize) =>
                    luckyConfirmed[prize.name].length > 0 && (
                      <div key={prize.name} className="mb-4 last:mb-0">
                        <h3 className="text-sm font-bold text-gray-600 mb-2">
                          {prize.name} ({luckyConfirmed[prize.name].length}/{prize.count}명)
                        </h3>
                        <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                          {luckyConfirmed[prize.name].map((p, i) => (
                            <div
                              key={i}
                              className="bg-primary-light text-primary rounded-lg p-2 text-center text-sm font-medium"
                            >
                              {i + 1}. {p.display}
                            </div>
                          ))}
                        </div>
                      </div>
                    )
                )}
              </div>
            )}
          </div>
        )}

        {/* Tab 3: Results */}
        {activeTab === 'results' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl shadow-md p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-lg font-bold text-primary">전체 당첨자 목록</h2>
                <div className="flex gap-2">
                  <button
                    onClick={loadResults}
                    className="text-sm bg-gray-100 hover:bg-gray-200 px-3 py-2 rounded-lg transition-colors"
                  >
                    새로고침
                  </button>
                  <button
                    onClick={copyResults}
                    className={`text-sm px-3 py-2 rounded-lg transition-colors ${
                      copied
                        ? 'bg-green-100 text-green-700'
                        : 'bg-primary text-white hover:bg-primary-dark'
                    }`}
                  >
                    {copied ? '복사됨!' : '결과 복사'}
                  </button>
                </div>
              </div>

              <div className="mb-6">
                <h3 className="text-base font-bold text-gray-700 mb-3 pb-2 border-b">
                  설문조사 경품추첨
                </h3>
                {allWinners.filter((w) => w.prize_type === 'survey').length === 0 ? (
                  <p className="text-gray-400 text-sm">아직 당첨자가 없습니다.</p>
                ) : (
                  <div className="space-y-2">
                    {allWinners
                      .filter((w) => w.prize_type === 'survey')
                      .map((w, i) => (
                        <div
                          key={w.id}
                          className="bg-primary-light text-primary rounded-lg p-3 text-sm"
                        >
                          <div className="font-bold">{i + 1}. {w.name}</div>
                        </div>
                      ))}
                  </div>
                )}
              </div>

              <div>
                <h3 className="text-base font-bold text-gray-700 mb-3 pb-2 border-b">
                  럭키드로우
                </h3>
                {allWinners.filter((w) => w.prize_type === 'lucky').length === 0 ? (
                  <p className="text-gray-400 text-sm">아직 당첨자가 없습니다.</p>
                ) : (
                  <>
                    {LUCKY_PRIZES.map((prize) => {
                      const winners = allWinners.filter(
                        (w) => w.prize_type === 'lucky' && w.prize === prize.name
                      )
                      if (winners.length === 0) return null
                      return (
                        <div key={prize.name} className="mb-4 last:mb-0">
                          <h4 className="text-sm font-bold text-gray-600 mb-2">{prize.name}</h4>
                          <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                            {winners.map((w, i) => (
                              <div
                                key={w.id}
                                className="bg-primary-light text-primary rounded-lg p-3 text-sm"
                              >
                                <div className="font-bold">{i + 1}. {w.name}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                      )
                    })}
                  </>
                )}
              </div>
            </div>
          </div>
        )}
      </main>

      <footer className="mt-auto py-4 text-center text-xs text-gray-400">
        주최: 교육부 | 주관: KRIVET 한국직업능력연구원
      </footer>
    </div>
  )
}
