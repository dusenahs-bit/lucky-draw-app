'use client'

import { useState, useRef, useCallback, useEffect } from 'react'
import { getSupabase, type Winner } from '../lib/supabase'

type Tab = 'survey' | 'lucky' | 'results'
type LuckyPrize = '논픽션 핸드크림' | '하이드로 텀블러' | 'TWG Tea'

const LUCKY_PRIZES: { name: LuckyPrize; count: number }[] = [
  { name: '논픽션 핸드크림', count: 5 },
  { name: '하이드로 텀블러', count: 3 },
  { name: 'TWG Tea', count: 3 },
]

function parseNames(text: string): string[] {
  return text
    .split(/[\n,]+/)
    .map((n) => n.trim())
    .filter((n) => n.length > 0)
}

function shuffle<T>(arr: T[]): T[] {
  const a = [...arr]
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[a[i], a[j]] = [a[j], a[i]]
  }
  return a
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<Tab>('survey')

  // Survey state
  const [surveyInput, setSurveyInput] = useState('')
  const [surveyCount, setSurveyCount] = useState(10)
  const [surveyDrawing, setSurveyDrawing] = useState(false)
  const [surveyDrumName, setSurveyDrumName] = useState('')
  const [surveyPending, setSurveyPending] = useState<string[]>([])
  const [surveyConfirmed, setSurveyConfirmed] = useState<string[]>([])
  const surveyTimerRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Lucky draw state
  const [luckyInput, setLuckyInput] = useState('')
  const [luckyPrizeTab, setLuckyPrizeTab] = useState<LuckyPrize>('논픽션 핸드크림')
  const [luckyDrawing, setLuckyDrawing] = useState(false)
  const [luckyDrumName, setLuckyDrumName] = useState('')
  const [luckyPending, setLuckyPending] = useState<string[]>([])
  const [luckyConfirmed, setLuckyConfirmed] = useState<Record<LuckyPrize, string[]>>({
    '논픽션 핸드크림': [],
    '하이드로 텀블러': [],
    'TWG Tea': [],
  })
  const luckyTimerRef = useRef<ReturnType<typeof setInterval> | null>(null)

  // Results state
  const [allWinners, setAllWinners] = useState<Winner[]>([])
  const [copied, setCopied] = useState(false)

  const getAllLuckyConfirmed = useCallback(() => {
    return Object.values(luckyConfirmed).flat()
  }, [luckyConfirmed])

  const getSurveyPool = useCallback(() => {
    const names = parseNames(surveyInput)
    const excluded = new Set(surveyConfirmed)
    return names.filter((n) => !excluded.has(n))
  }, [surveyInput, surveyConfirmed])

  const getLuckyPool = useCallback(() => {
    const names = parseNames(luckyInput)
    const excluded = new Set(getAllLuckyConfirmed())
    return names.filter((n) => !excluded.has(n))
  }, [luckyInput, getAllLuckyConfirmed])

  const currentPrizeConfig = LUCKY_PRIZES.find((p) => p.name === luckyPrizeTab)!
  const currentPrizeRemaining = currentPrizeConfig.count - luckyConfirmed[luckyPrizeTab].length

  const startDrumRoll = (
    pool: string[],
    count: number,
    setDrumName: (n: string) => void,
    setDrawing: (b: boolean) => void,
    setPending: (names: string[]) => void,
    timerRef: React.MutableRefObject<ReturnType<typeof setInterval> | null>
  ) => {
    if (pool.length === 0) return
    const drawCount = Math.min(count, pool.length)
    setDrawing(true)
    setPending([])
    setDrumName(pool[Math.floor(Math.random() * pool.length)])

    let tick = 0
    const totalTicks = 30

    const run = () => {
      tick++
      const randomIdx = Math.floor(Math.random() * pool.length)
      setDrumName(pool[randomIdx])

      if (tick >= totalTicks) {
        if (timerRef.current) clearInterval(timerRef.current)
        const winners = shuffle(pool).slice(0, drawCount)
        setPending(winners)
        setDrumName('')
        setDrawing(false)
      }
    }

    timerRef.current = setInterval(run, 100)
  }

  const handleSurveyDraw = () => {
    const pool = getSurveyPool()
    setSurveyPending([])
    startDrumRoll(pool, surveyCount, setSurveyDrumName, setSurveyDrawing, setSurveyPending, surveyTimerRef)
  }

  const handleSurveyConfirm = async () => {
    const newConfirmed = [...surveyConfirmed, ...surveyPending]
    setSurveyConfirmed(newConfirmed)

    const rows = surveyPending.map((name) => ({
      name,
      prize: '배달의민족 상품권 5만원권',
      prize_type: 'survey' as const,
    }))

    await getSupabase().from('winners').insert(rows)
    setSurveyPending([])
  }

  const handleSurveyRedraw = () => {
    const pool = getSurveyPool().filter((n) => !surveyPending.includes(n))
    setSurveyPending([])
    startDrumRoll(pool, surveyCount, setSurveyDrumName, setSurveyDrawing, setSurveyPending, surveyTimerRef)
  }

  const handleSurveyRedrawOne = (idx: number) => {
    const excluded = new Set([...surveyConfirmed, ...surveyPending])
    const pool = parseNames(surveyInput).filter((n) => !excluded.has(n))
    if (pool.length === 0) return
    const replacement = shuffle(pool)[0]
    setSurveyPending((prev) => prev.map((n, i) => (i === idx ? replacement : n)))
  }

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

    const rows = luckyPending.map((name) => ({
      name,
      prize: luckyPrizeTab,
      prize_type: 'lucky' as const,
    }))

    await getSupabase().from('winners').insert(rows)
    setLuckyPending([])
  }

  const handleLuckyRedrawOne = (idx: number) => {
    const allConfirmed = getAllLuckyConfirmed()
    const excluded = new Set([...allConfirmed, ...luckyPending])
    const pool = parseNames(luckyInput).filter((n) => !excluded.has(n))
    if (pool.length === 0) return
    const replacement = shuffle(pool)[0]
    setLuckyPending((prev) => prev.map((n, i) => (i === idx ? replacement : n)))
  }

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

  const handleFileUpload = (
    e: React.ChangeEvent<HTMLInputElement>,
    setInput: (v: string) => void
  ) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (ev) => {
      const text = ev.target?.result as string
      setInput(text)
    }
    reader.readAsText(file)
  }

  const copyResults = () => {
    const surveyWinners = allWinners.filter((w) => w.prize_type === 'survey')
    const luckyWinners = allWinners.filter((w) => w.prize_type === 'lucky')

    let text = '=== 설문조사 경품추첨 당첨자 ===\n'
    surveyWinners.forEach((w, i) => {
      text += `${i + 1}. ${w.name} - ${w.prize}\n`
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

  const tabs: { key: Tab; label: string }[] = [
    { key: 'survey', label: '설문조사 추첨' },
    { key: 'lucky', label: '럭키드로우' },
    { key: 'results', label: '최종 결과' },
  ]

  return (
    <div className="min-h-screen bg-gradient-to-b from-[#e8f4fd] to-white">
      {/* Header */}
      <header className="bg-primary text-white py-4 px-4 shadow-lg">
        <h1 className="text-center text-xl md:text-2xl font-bold">
          2026년 국가·지역진로교육센터 및 진로체험지원센터 워크숍
        </h1>
        <p className="text-center text-sm md:text-base mt-1 opacity-90">경품추첨 이벤트</p>
      </header>

      {/* Tab navigation */}
      <nav className="flex border-b border-gray-200 bg-white sticky top-0 z-10">
        {tabs.map((tab) => (
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
            <div className="bg-white rounded-xl shadow-md p-6">
              <h2 className="text-lg font-bold text-primary mb-4">참가자 명단 입력</h2>
              <textarea
                className="w-full h-40 border border-gray-300 rounded-lg p-3 text-sm focus:outline-none focus:ring-2 focus:ring-primary/50 resize-none"
                placeholder="이름을 줄바꿈 또는 쉼표로 구분하여 입력하세요..."
                value={surveyInput}
                onChange={(e) => setSurveyInput(e.target.value)}
              />
              <div className="mt-3 flex items-center gap-3">
                <label className="text-sm text-gray-600 cursor-pointer bg-gray-100 hover:bg-gray-200 px-3 py-2 rounded-lg transition-colors">
                  파일 업로드 (txt/csv)
                  <input
                    type="file"
                    accept=".txt,.csv"
                    className="hidden"
                    onChange={(e) => handleFileUpload(e, setSurveyInput)}
                  />
                </label>
                <span className="text-sm text-gray-500">
                  참가자 {getSurveyPool().length}명 (전체 {parseNames(surveyInput).length}명)
                </span>
              </div>
            </div>

            <div className="bg-white rounded-xl shadow-md p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-lg font-bold text-primary">추첨 인원</h2>
                <span className="text-2xl font-bold text-primary">{surveyCount}명</span>
              </div>
              <input
                type="range"
                min={1}
                max={Math.max(1, getSurveyPool().length)}
                value={surveyCount}
                onChange={(e) => setSurveyCount(Number(e.target.value))}
                className="w-full accent-[#1a6fb5]"
              />
              <div className="flex justify-between text-xs text-gray-400 mt-1">
                <span>1명</span>
                <span>{Math.max(1, getSurveyPool().length)}명</span>
              </div>
            </div>

            {/* Drum roll area */}
            <div className="bg-white rounded-xl shadow-md p-6 text-center">
              {surveyDrawing && (
                <div className="mb-4">
                  <div className="text-4xl md:text-5xl font-bold text-primary drumroll-active py-8">
                    {surveyDrumName || '...'}
                  </div>
                  <p className="text-gray-500 mt-2">추첨 중...</p>
                </div>
              )}

              {!surveyDrawing && surveyPending.length === 0 && (
                <button
                  onClick={handleSurveyDraw}
                  disabled={getSurveyPool().length === 0}
                  className="bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-lg font-bold py-4 px-12 rounded-full transition-colors shadow-lg hover:shadow-xl"
                >
                  추첨 시작
                </button>
              )}

              {!surveyDrawing && surveyPending.length > 0 && (
                <div>
                  <h3 className="text-lg font-bold text-gray-700 mb-4">당첨자 발표!</h3>
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-6">
                    {surveyPending.map((name, i) => (
                      <div
                        key={i}
                        className="winner-reveal bg-gradient-to-br from-primary to-primary-dark text-white rounded-lg p-3 shadow-md"
                        style={{ animationDelay: `${i * 0.1}s` }}
                      >
                        <div className="text-lg font-bold">{name}</div>
                        <button
                          onClick={() => handleSurveyRedrawOne(i)}
                          className="text-xs mt-1 opacity-75 hover:opacity-100 underline"
                        >
                          재추첨
                        </button>
                      </div>
                    ))}
                  </div>
                  <div className="flex gap-3 justify-center">
                    <button
                      onClick={handleSurveyConfirm}
                      className="bg-green-500 hover:bg-green-600 text-white font-bold py-3 px-8 rounded-full transition-colors"
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

            {/* Confirmed winners */}
            {surveyConfirmed.length > 0 && (
              <div className="bg-white rounded-xl shadow-md p-6">
                <h2 className="text-lg font-bold text-primary mb-4">
                  확정 당첨자 ({surveyConfirmed.length}명)
                </h2>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                  {surveyConfirmed.map((name, i) => (
                    <div
                      key={i}
                      className="bg-primary-light text-primary rounded-lg p-2 text-center text-sm font-medium"
                    >
                      {i + 1}. {name}
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
            <div className="bg-white rounded-xl shadow-md p-6">
              <h2 className="text-lg font-bold text-primary mb-4">참가자 명단 입력</h2>
              <textarea
                className="w-full h-40 border border-gray-300 rounded-lg p-3 text-sm focus:outline-none focus:ring-2 focus:ring-primary/50 resize-none"
                placeholder="이름을 줄바꿈 또는 쉼표로 구분하여 입력하세요..."
                value={luckyInput}
                onChange={(e) => setLuckyInput(e.target.value)}
              />
              <div className="mt-3 flex items-center gap-3">
                <label className="text-sm text-gray-600 cursor-pointer bg-gray-100 hover:bg-gray-200 px-3 py-2 rounded-lg transition-colors">
                  파일 업로드 (txt/csv)
                  <input
                    type="file"
                    accept=".txt,.csv"
                    className="hidden"
                    onChange={(e) => handleFileUpload(e, setLuckyInput)}
                  />
                </label>
                <span className="text-sm text-gray-500">
                  참가자 {getLuckyPool().length}명 (전체 {parseNames(luckyInput).length}명)
                </span>
              </div>
            </div>

            {/* Prize tabs */}
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

              <div className="p-6 text-center">
                {luckyDrawing && (
                  <div className="mb-4">
                    <div className="text-4xl md:text-5xl font-bold text-primary drumroll-active py-8">
                      {luckyDrumName || '...'}
                    </div>
                    <p className="text-gray-500 mt-2">추첨 중...</p>
                  </div>
                )}

                {!luckyDrawing && luckyPending.length === 0 && (
                  <div>
                    <p className="text-gray-500 mb-4">
                      {luckyPrizeTab} — 남은 추첨: {currentPrizeRemaining}명
                    </p>
                    <button
                      onClick={handleLuckyDraw}
                      disabled={currentPrizeRemaining <= 0 || getLuckyPool().length === 0}
                      className="bg-primary hover:bg-primary-dark disabled:bg-gray-300 text-white text-lg font-bold py-4 px-12 rounded-full transition-colors shadow-lg hover:shadow-xl"
                    >
                      {currentPrizeRemaining <= 0 ? '추첨 완료' : '추첨 시작'}
                    </button>
                  </div>
                )}

                {!luckyDrawing && luckyPending.length > 0 && (
                  <div>
                    <h3 className="text-lg font-bold text-gray-700 mb-4">
                      {luckyPrizeTab} 당첨자!
                    </h3>
                    <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-6">
                      {luckyPending.map((name, i) => (
                        <div
                          key={i}
                          className="winner-reveal bg-gradient-to-br from-primary to-primary-dark text-white rounded-lg p-3 shadow-md"
                          style={{ animationDelay: `${i * 0.1}s` }}
                        >
                          <div className="text-lg font-bold">{name}</div>
                          <button
                            onClick={() => handleLuckyRedrawOne(i)}
                            className="text-xs mt-1 opacity-75 hover:opacity-100 underline"
                          >
                            재추첨
                          </button>
                        </div>
                      ))}
                    </div>
                    <div className="flex gap-3 justify-center">
                      <button
                        onClick={handleLuckyConfirm}
                        className="bg-green-500 hover:bg-green-600 text-white font-bold py-3 px-8 rounded-full transition-colors"
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

            {/* Lucky confirmed list */}
            {getAllLuckyConfirmed().length > 0 && (
              <div className="bg-white rounded-xl shadow-md p-6">
                <h2 className="text-lg font-bold text-primary mb-4">럭키드로우 당첨자</h2>
                {LUCKY_PRIZES.map(
                  (prize) =>
                    luckyConfirmed[prize.name].length > 0 && (
                      <div key={prize.name} className="mb-4 last:mb-0">
                        <h3 className="text-sm font-bold text-gray-600 mb-2">
                          {prize.name} ({luckyConfirmed[prize.name].length}/{prize.count}명)
                        </h3>
                        <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                          {luckyConfirmed[prize.name].map((name, i) => (
                            <div
                              key={i}
                              className="bg-primary-light text-primary rounded-lg p-2 text-center text-sm font-medium"
                            >
                              {i + 1}. {name}
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

              {/* Survey winners */}
              <div className="mb-6">
                <h3 className="text-base font-bold text-gray-700 mb-3 pb-2 border-b">
                  설문조사 경품추첨
                </h3>
                {allWinners.filter((w) => w.prize_type === 'survey').length === 0 ? (
                  <p className="text-gray-400 text-sm">아직 당첨자가 없습니다.</p>
                ) : (
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                    {allWinners
                      .filter((w) => w.prize_type === 'survey')
                      .map((w, i) => (
                        <div
                          key={w.id}
                          className="bg-primary-light text-primary rounded-lg p-3 text-sm"
                        >
                          <div className="font-bold">
                            {i + 1}. {w.name}
                          </div>
                          <div className="text-xs opacity-75">{w.prize}</div>
                        </div>
                      ))}
                  </div>
                )}
              </div>

              {/* Lucky draw winners */}
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
                                <div className="font-bold">
                                  {i + 1}. {w.name}
                                </div>
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
