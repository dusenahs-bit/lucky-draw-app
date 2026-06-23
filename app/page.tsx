'use client'

import { useState, useRef, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { getSupabase, type Winner } from '../lib/supabase'

type Mode = 'setup' | 'draw'
type DrawPage = 'home' | 'survey' | 'lucky' | 'results'
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
      return { key: idInfo, drumName: idInfo, display: `${type} | ${region} | ${idInfo}` }
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
  const [drawPage, setDrawPage] = useState<DrawPage>('home')

  // Survey state
  const [surveyParticipants, setSurveyParticipants] = useState<Participant[]>([])
  const [surveyFileName, setSurveyFileName] = useState('')
  const [surveyCount, setSurveyCount] = useState(10)
  const [surveyDrawing, setSurveyDrawing] = useState(false)
  const [surveyDrumName, setSurveyDrumName] = useState('')
  const [surveyPending, setSurveyPending] = useState<Participant[]>([])
  const [surveyPendingConfirmed, setSurveyPendingConfirmed] = useState<Set<string>>(new Set())
  const [surveyConfirmed, setSurveyConfirmed] = useState<Participant[]>([])
  const surveyTimerRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Lucky draw state
  const [luckyParticipants, setLuckyParticipants] = useState<Participant[]>([])
  const [luckyFileName, setLuckyFileName] = useState('')
  const [luckyPrizeTab, setLuckyPrizeTab] = useState<LuckyPrize>('논픽션 핸드크림')
  const [luckyDrawing, setLuckyDrawing] = useState(false)
  const [luckyDrumName, setLuckyDrumName] = useState('')
  const [luckyPending, setLuckyPending] = useState<Participant[]>([])
  const [luckyPendingConfirmed, setLuckyPendingConfirmed] = useState<Set<string>>(new Set())
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
        setPending(shuffle(pool).slice(0, drawCount))
        setDrumNameFn('')
        setDrawing(false)
      }
    }
    timerRef.current = setInterval(run, 100)
  }

  // Survey handlers
  const handleSurveyDraw = () => {
    setSurveyPending([])
    setSurveyPendingConfirmed(new Set())
    startDrumRoll(getSurveyPool(), surveyCount, setSurveyDrumName, setSurveyDrawing, setSurveyPending, surveyTimerRef)
  }

  const handleSurveyConfirmOne = async (p: Participant) => {
    setSurveyConfirmed((prev) => [...prev, p])
    setSurveyPendingConfirmed((prev) => new Set(prev).add(p.key))
    await getSupabase().from('winners').insert({
      name: p.display,
      prize: '배달의민족 상품권 5만원권',
      prize_type: 'survey',
    })
  }

  const handleSurveyRedrawOne = (idx: number) => {
    const excludedKeys = new Set([
      ...surveyConfirmed.map((p) => p.key),
      ...surveyPending.map((p) => p.key),
    ])
    const pool = surveyParticipants.filter((p) => !excludedKeys.has(p.key))
    if (pool.length === 0) return
    setSurveyPending((prev) => prev.map((p, i) => (i === idx ? shuffle(pool)[0] : p)))
  }

  const handleSurveyRedrawUnconfirmed = () => {
    const unconfirmed = surveyPending.filter((p) => !surveyPendingConfirmed.has(p.key))
    const count = unconfirmed.length
    if (count === 0) return
    const excludedKeys = new Set([
      ...surveyConfirmed.map((p) => p.key),
      ...surveyPending.filter((p) => surveyPendingConfirmed.has(p.key)).map((p) => p.key),
    ])
    const pool = surveyParticipants.filter((p) => !excludedKeys.has(p.key))
    const newWinners = shuffle(pool).slice(0, Math.min(count, pool.length))
    let newIdx = 0
    setSurveyPending((prev) =>
      prev.map((p) => {
        if (!surveyPendingConfirmed.has(p.key) && newIdx < newWinners.length) {
          return newWinners[newIdx++]
        }
        return p
      })
    )
  }

  const surveyAllConfirmed =
    surveyPending.length > 0 && surveyPending.every((p) => surveyPendingConfirmed.has(p.key))

  const handleSurveyFinishBatch = () => {
    setSurveyPending([])
    setSurveyPendingConfirmed(new Set())
  }

  // Lucky draw handlers
  const handleLuckyDraw = () => {
    setLuckyPending([])
    setLuckyPendingConfirmed(new Set())
    startDrumRoll(getLuckyPool(), currentPrizeRemaining, setLuckyDrumName, setLuckyDrawing, setLuckyPending, luckyTimerRef)
  }

  const handleLuckyConfirmOne = async (p: Participant) => {
    setLuckyConfirmed((prev) => ({
      ...prev,
      [luckyPrizeTab]: [...prev[luckyPrizeTab], p],
    }))
    setLuckyPendingConfirmed((prev) => new Set(prev).add(p.key))
    await getSupabase().from('winners').insert({
      name: p.display,
      prize: luckyPrizeTab,
      prize_type: 'lucky',
    })
  }

  const handleLuckyRedrawOne = (idx: number) => {
    const allConfirmedKeys = getAllLuckyConfirmedKeys()
    const pendingKeys = new Set(luckyPending.map((p) => p.key))
    const pool = luckyParticipants.filter(
      (p) => !allConfirmedKeys.has(p.key) && !pendingKeys.has(p.key)
    )
    if (pool.length === 0) return
    setLuckyPending((prev) => prev.map((p, i) => (i === idx ? shuffle(pool)[0] : p)))
  }

  const handleLuckyRedrawUnconfirmed = () => {
    const unconfirmed = luckyPending.filter((p) => !luckyPendingConfirmed.has(p.key))
    const count = unconfirmed.length
    if (count === 0) return
    const allConfirmedKeys = getAllLuckyConfirmedKeys()
    const keepKeys = new Set(luckyPending.filter((p) => luckyPendingConfirmed.has(p.key)).map((p) => p.key))
    const pool = luckyParticipants.filter(
      (p) => !allConfirmedKeys.has(p.key) && !keepKeys.has(p.key)
    )
    const newWinners = shuffle(pool).slice(0, Math.min(count, pool.length))
    let newIdx = 0
    setLuckyPending((prev) =>
      prev.map((p) => {
        if (!luckyPendingConfirmed.has(p.key) && newIdx < newWinners.length) {
          return newWinners[newIdx++]
        }
        return p
      })
    )
  }

  const luckyAllConfirmed =
    luckyPending.length > 0 && luckyPending.every((p) => luckyPendingConfirmed.has(p.key))

  const handleLuckyFinishBatch = () => {
    setLuckyPending([])
    setLuckyPendingConfirmed(new Set())
  }

  // File upload handlers
  const handleSurveyFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setSurveyFileName(file.name)
    const reader = new FileReader()
    reader.onload = (ev) => {
      setSurveyParticipants(parseSurveyExcel(new Uint8Array(ev.target?.result as ArrayBuffer)))
    }
    reader.readAsArrayBuffer(file)
  }

  const handleLuckyFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setLuckyFileName(file.name)
    const reader = new FileReader()
    reader.onload = (ev) => {
      setLuckyParticipants(parseLuckyExcel(new Uint8Array(ev.target?.result as ArrayBuffer)))
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
    if (drawPage === 'results') loadResults()
  }, [drawPage])

  const copyResults = () => {
    const surveyWinners = allWinners.filter((w) => w.prize_type === 'survey')
    const luckyWinners = allWinners.filter((w) => w.prize_type === 'lucky')
    let text = '=== 설문조사 경품추첨 당첨자 ===\n'
    surveyWinners.forEach((w, i) => { text += `${i + 1}. ${w.name}\n` })
    text += '\n=== 럭키드로우 당첨자 ===\n'
    const grouped: Record<string, string[]> = {}
    luckyWinners.forEach((w) => {
      if (!grouped[w.prize]) grouped[w.prize] = []
      grouped[w.prize].push(w.name)
    })
    Object.entries(grouped).forEach(([prize, names]) => {
      text += `\n[${prize}]\n`
      names.forEach((n, i) => { text += `${i + 1}. ${n}\n` })
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
          <div className="bg-white rounded-xl shadow-md p-6">
            <h2 className="text-lg font-bold text-gray-800 mb-2">1. 설문조사 추첨 명단</h2>
            <p className="text-sm text-gray-500 mb-4">
              엑셀 형식: A열(센터유형), B열(지역), C열(아이디+연락처 뒷번호)
            </p>
            <div className="flex items-center gap-3">
              <label className="text-sm text-white cursor-pointer bg-gray-700 hover:bg-gray-800 px-4 py-2 rounded-lg transition-colors font-medium">
                파일 선택
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleSurveyFileUpload} />
              </label>
              {surveyFileName && <span className="text-sm text-gray-600">{surveyFileName}</span>}
            </div>
            {surveyParticipants.length > 0 && (
              <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg">
                <span className="text-sm font-medium text-green-700">{surveyParticipants.length}명 로드 완료</span>
                <div className="mt-2 max-h-40 overflow-y-auto text-xs text-gray-600 space-y-1">
                  {surveyParticipants.map((p, i) => (<div key={i} className="py-0.5">{i + 1}. {p.display}</div>))}
                </div>
              </div>
            )}
            <div className="mt-4">
              <label className="text-sm font-medium text-gray-700">추첨 인원</label>
              <div className="flex items-center gap-3 mt-1">
                <input type="number" min={1} max={Math.max(1, surveyParticipants.length)} value={surveyCount}
                  onChange={(e) => setSurveyCount(Number(e.target.value))}
                  className="w-20 border border-gray-300 rounded-lg px-3 py-2 text-sm" />
                <span className="text-sm text-gray-500">명</span>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-md p-6">
            <h2 className="text-lg font-bold text-gray-800 mb-2">2. 럭키드로우 명단</h2>
            <p className="text-sm text-gray-500 mb-4">
              엑셀 형식: B열(이름), C열(연락처) — 당첨 시 이름 + 연락처 뒷4자리 표시
            </p>
            <div className="flex items-center gap-3">
              <label className="text-sm text-white cursor-pointer bg-gray-700 hover:bg-gray-800 px-4 py-2 rounded-lg transition-colors font-medium">
                파일 선택
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleLuckyFileUpload} />
              </label>
              {luckyFileName && <span className="text-sm text-gray-600">{luckyFileName}</span>}
            </div>
            {luckyParticipants.length > 0 && (
              <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg">
                <span className="text-sm font-medium text-green-700">{luckyParticipants.length}명 로드 완료</span>
                <div className="mt-2 max-h-40 overflow-y-auto text-xs text-gray-600 space-y-1">
                  {luckyParticipants.map((p, i) => (<div key={i} className="py-0.5">{i + 1}. {p.display}</div>))}
                </div>
              </div>
            )}
          </div>

          <button onClick={() => { setMode('draw'); setDrawPage('home') }} disabled={!canStartDraw}
            className="w-full bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-lg font-bold py-4 rounded-xl transition-colors shadow-lg">
            {canStartDraw ? '추첨 화면으로 전환' : '명단을 먼저 업로드하세요'}
          </button>
        </main>
      </div>
    )
  }

  // ==================== DRAW MODE ====================
  return (
    <div className="min-h-screen bg-gradient-to-b from-[#e8f4fd] to-white">
      <header className="bg-primary text-white py-5 px-4 shadow-lg relative">
        <h1 className="text-center text-2xl md:text-3xl font-bold">경품추첨</h1>
        <button onClick={() => setMode('setup')}
          className="absolute top-5 right-4 bg-white/20 hover:bg-white/40 text-white text-sm px-3 py-1 rounded-lg transition-colors" title="관리자 설정">
          관리자
        </button>
      </header>

      <main className="max-w-3xl mx-auto p-4 md:p-6">
        {/* HOME: Draw selection */}
        {drawPage === 'home' && (
          <div className="space-y-4 mt-8">
            <button onClick={() => setDrawPage('survey')}
              className="w-full bg-white hover:bg-primary-light rounded-2xl shadow-lg p-8 text-left transition-colors group">
              <h2 className="text-2xl font-bold text-primary group-hover:text-primary-dark">설문조사 경품추첨</h2>
              <p className="text-gray-500 mt-2">배달의민족 상품권 5만원권 · {surveyCount}명</p>
              {surveyConfirmed.length > 0 && (
                <p className="text-sm text-green-600 mt-1">확정 {surveyConfirmed.length}명</p>
              )}
            </button>
            <button onClick={() => setDrawPage('lucky')}
              className="w-full bg-white hover:bg-primary-light rounded-2xl shadow-lg p-8 text-left transition-colors group">
              <h2 className="text-2xl font-bold text-primary group-hover:text-primary-dark">럭키드로우</h2>
              <p className="text-gray-500 mt-2">논픽션 핸드크림 5명 · 하이드로 텀블러 3명 · TWG Tea 3명</p>
              {Object.values(luckyConfirmed).flat().length > 0 && (
                <p className="text-sm text-green-600 mt-1">확정 {Object.values(luckyConfirmed).flat().length}명</p>
              )}
            </button>
            <button onClick={() => setDrawPage('results')}
              className="w-full bg-white hover:bg-gray-50 rounded-2xl shadow-lg p-8 text-left transition-colors group">
              <h2 className="text-2xl font-bold text-gray-600 group-hover:text-gray-800">최종 결과</h2>
              <p className="text-gray-400 mt-2">전체 당첨자 조회 및 복사</p>
            </button>
          </div>
        )}

        {/* SURVEY DRAW */}
        {drawPage === 'survey' && (
          <div className="space-y-6">
            <button onClick={() => setDrawPage('home')} className="text-sm text-gray-400 hover:text-gray-600">
              ← 돌아가기
            </button>

            <div className="bg-white rounded-xl shadow-md p-6 text-center min-h-[350px] flex flex-col items-center justify-center">
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
                  <h2 className="text-primary font-bold text-2xl mb-8">
                    배달의민족 상품권 5만원권
                  </h2>
                  <button onClick={handleSurveyDraw} disabled={getSurveyPool().length === 0}
                    className="bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-xl font-bold py-5 px-16 rounded-full transition-colors shadow-lg hover:shadow-xl">
                    추첨 시작 ({surveyCount}명)
                  </button>
                </div>
              )}

              {!surveyDrawing && surveyPending.length > 0 && (
                <div className="w-full">
                  <h3 className="text-2xl font-bold text-gray-700 mb-6">당첨을 축하합니다!</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-8">
                    {surveyPending.map((p, i) => {
                      const isConfirmed = surveyPendingConfirmed.has(p.key)
                      return (
                        <div key={p.key}
                          className={`winner-reveal rounded-xl p-4 shadow-lg text-left transition-all ${
                            isConfirmed
                              ? 'bg-green-500 text-white'
                              : 'bg-gradient-to-br from-primary to-primary-dark text-white'
                          }`}
                          style={{ animationDelay: `${i * 0.1}s` }}>
                          <div className="text-base font-bold">{p.display}</div>
                          {isConfirmed ? (
                            <span className="text-xs mt-2 inline-block opacity-80">확인 완료</span>
                          ) : (
                            <div className="flex gap-2 mt-3">
                              <button onClick={() => handleSurveyConfirmOne(p)}
                                className="text-xs bg-white/20 hover:bg-white/30 px-3 py-1 rounded-full transition-colors">
                                확인
                              </button>
                              <button onClick={() => handleSurveyRedrawOne(i)}
                                className="text-xs bg-white/10 hover:bg-white/20 px-3 py-1 rounded-full transition-colors">
                                재추첨
                              </button>
                            </div>
                          )}
                        </div>
                      )
                    })}
                  </div>
                  <div className="flex gap-3 justify-center">
                    {surveyAllConfirmed ? (
                      <button onClick={handleSurveyFinishBatch}
                        className="bg-primary hover:bg-primary-dark text-white font-bold py-3 px-8 rounded-full transition-colors text-lg">
                        다음 추첨
                      </button>
                    ) : (
                      <button onClick={handleSurveyRedrawUnconfirmed}
                        className="bg-gray-400 hover:bg-gray-500 text-white font-bold py-3 px-8 rounded-full transition-colors">
                        미확인자 재추첨 ({surveyPending.filter((p) => !surveyPendingConfirmed.has(p.key)).length}명)
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>

            {surveyConfirmed.length > 0 && (
              <div className="bg-white rounded-xl shadow-md p-6">
                <h2 className="text-lg font-bold text-primary mb-4">확정 당첨자 ({surveyConfirmed.length}명)</h2>
                <div className="space-y-2">
                  {surveyConfirmed.map((p, i) => (
                    <div key={i} className="bg-primary-light text-primary rounded-lg p-3 text-sm font-medium">
                      {i + 1}. {p.display}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* LUCKY DRAW */}
        {drawPage === 'lucky' && (
          <div className="space-y-6">
            <button onClick={() => setDrawPage('home')} className="text-sm text-gray-400 hover:text-gray-600">
              ← 돌아가기
            </button>

            <div className="bg-white rounded-xl shadow-md overflow-hidden">
              <div className="flex border-b">
                {LUCKY_PRIZES.map((prize) => (
                  <button key={prize.name}
                    onClick={() => { setLuckyPrizeTab(prize.name); setLuckyPending([]); setLuckyPendingConfirmed(new Set()) }}
                    className={`flex-1 py-3 px-2 text-sm font-medium transition-colors ${
                      luckyPrizeTab === prize.name
                        ? 'text-primary border-b-2 border-primary bg-primary-light/50'
                        : 'text-gray-500 hover:text-gray-700'
                    }`}>
                    {prize.name}<br />
                    <span className="text-xs">({luckyConfirmed[prize.name].length}/{prize.count}명)</span>
                  </button>
                ))}
              </div>

              <div className="p-6 text-center min-h-[350px] flex flex-col items-center justify-center">
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
                    <h2 className="text-primary font-bold text-2xl mb-8">{luckyPrizeTab}</h2>
                    <button onClick={handleLuckyDraw}
                      disabled={currentPrizeRemaining <= 0 || getLuckyPool().length === 0}
                      className="bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-xl font-bold py-5 px-16 rounded-full transition-colors shadow-lg hover:shadow-xl">
                      {currentPrizeRemaining <= 0 ? '추첨 완료' : `추첨 시작 (${currentPrizeRemaining}명)`}
                    </button>
                  </div>
                )}

                {!luckyDrawing && luckyPending.length > 0 && (
                  <div className="w-full">
                    <h3 className="text-2xl font-bold text-gray-700 mb-6">{luckyPrizeTab} 당첨!</h3>
                    <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-8">
                      {luckyPending.map((p, i) => {
                        const isConfirmed = luckyPendingConfirmed.has(p.key)
                        return (
                          <div key={p.key}
                            className={`winner-reveal rounded-xl p-4 shadow-lg transition-all ${
                              isConfirmed
                                ? 'bg-green-500 text-white'
                                : 'bg-gradient-to-br from-primary to-primary-dark text-white'
                            }`}
                            style={{ animationDelay: `${i * 0.1}s` }}>
                            <div className="text-lg font-bold">{p.display}</div>
                            {isConfirmed ? (
                              <span className="text-xs mt-2 inline-block opacity-80">확인 완료</span>
                            ) : (
                              <div className="flex gap-2 mt-3">
                                <button onClick={() => handleLuckyConfirmOne(p)}
                                  className="text-xs bg-white/20 hover:bg-white/30 px-3 py-1 rounded-full transition-colors">
                                  확인
                                </button>
                                <button onClick={() => handleLuckyRedrawOne(i)}
                                  className="text-xs bg-white/10 hover:bg-white/20 px-3 py-1 rounded-full transition-colors">
                                  재추첨
                                </button>
                              </div>
                            )}
                          </div>
                        )
                      })}
                    </div>
                    <div className="flex gap-3 justify-center">
                      {luckyAllConfirmed ? (
                        <button onClick={handleLuckyFinishBatch}
                          className="bg-primary hover:bg-primary-dark text-white font-bold py-3 px-8 rounded-full transition-colors text-lg">
                          다음 추첨
                        </button>
                      ) : (
                        <button onClick={handleLuckyRedrawUnconfirmed}
                          className="bg-gray-400 hover:bg-gray-500 text-white font-bold py-3 px-8 rounded-full transition-colors">
                          미확인자 재추첨 ({luckyPending.filter((p) => !luckyPendingConfirmed.has(p.key)).length}명)
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>

            {Object.values(luckyConfirmed).flat().length > 0 && (
              <div className="bg-white rounded-xl shadow-md p-6">
                <h2 className="text-lg font-bold text-primary mb-4">럭키드로우 당첨자</h2>
                {LUCKY_PRIZES.map((prize) =>
                  luckyConfirmed[prize.name].length > 0 && (
                    <div key={prize.name} className="mb-4 last:mb-0">
                      <h3 className="text-sm font-bold text-gray-600 mb-2">
                        {prize.name} ({luckyConfirmed[prize.name].length}/{prize.count}명)
                      </h3>
                      <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                        {luckyConfirmed[prize.name].map((p, i) => (
                          <div key={i} className="bg-primary-light text-primary rounded-lg p-2 text-center text-sm font-medium">
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

        {/* RESULTS */}
        {drawPage === 'results' && (
          <div className="space-y-6">
            <button onClick={() => setDrawPage('home')} className="text-sm text-gray-400 hover:text-gray-600">
              ← 돌아가기
            </button>
            <div className="bg-white rounded-xl shadow-md p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-lg font-bold text-primary">전체 당첨자 목록</h2>
                <div className="flex gap-2">
                  <button onClick={loadResults} className="text-sm bg-gray-100 hover:bg-gray-200 px-3 py-2 rounded-lg transition-colors">
                    새로고침
                  </button>
                  <button onClick={copyResults}
                    className={`text-sm px-3 py-2 rounded-lg transition-colors ${
                      copied ? 'bg-green-100 text-green-700' : 'bg-primary text-white hover:bg-primary-dark'
                    }`}>
                    {copied ? '복사됨!' : '결과 복사'}
                  </button>
                </div>
              </div>
              <div className="mb-6">
                <h3 className="text-base font-bold text-gray-700 mb-3 pb-2 border-b">설문조사 경품추첨</h3>
                {allWinners.filter((w) => w.prize_type === 'survey').length === 0 ? (
                  <p className="text-gray-400 text-sm">아직 당첨자가 없습니다.</p>
                ) : (
                  <div className="space-y-2">
                    {allWinners.filter((w) => w.prize_type === 'survey').map((w, i) => (
                      <div key={w.id} className="bg-primary-light text-primary rounded-lg p-3 text-sm">
                        <div className="font-bold">{i + 1}. {w.name}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-700 mb-3 pb-2 border-b">럭키드로우</h3>
                {allWinners.filter((w) => w.prize_type === 'lucky').length === 0 ? (
                  <p className="text-gray-400 text-sm">아직 당첨자가 없습니다.</p>
                ) : (
                  <>
                    {LUCKY_PRIZES.map((prize) => {
                      const winners = allWinners.filter((w) => w.prize_type === 'lucky' && w.prize === prize.name)
                      if (winners.length === 0) return null
                      return (
                        <div key={prize.name} className="mb-4 last:mb-0">
                          <h4 className="text-sm font-bold text-gray-600 mb-2">{prize.name}</h4>
                          <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                            {winners.map((w, i) => (
                              <div key={w.id} className="bg-primary-light text-primary rounded-lg p-3 text-sm">
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
