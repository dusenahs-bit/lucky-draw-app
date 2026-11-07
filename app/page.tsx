'use client'

import { useState, useRef, useCallback, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { getSupabase } from '../lib/supabase'

// ===== TYPES =====
interface Participant {
  key: string
  name: string
  display: string
}

interface Prize {
  id: number
  name: string
  count: number
  drawer: string
  noRedraw?: boolean  // true: 재추첨 없이 전체 즉시 확정 (모바일 발송용)
}

type Mode = 'setup' | 'draw'
type DrawPage = 'home' | 'prize' | 'results'
type DrawPhase = 'idle' | 'countdown' | 'drumroll' | 'reveal'

// ===== CONSTANTS =====
const DEFAULT_PRIZES: Prize[] = [
  { id: 1, name: '루이스폴센 판텔라 160 포터블 램프 V3', count: 2, drawer: '김미희 대구경북 분회장' },
  { id: 2, name: '보스 QC 울트라 헤드폰', count: 2, drawer: '김현정 호남분회장' },
  { id: 3, name: '로보락 S8 Pro Plus 로봇청소기', count: 1, drawer: '양귀자 부산경남분회장' },
  { id: 4, name: '애플 아이패드 에어11(M4 256GB)', count: 1, drawer: '김미형 회장' },
  { id: 5, name: 'LG 스탠마이미2', count: 1, drawer: '주미자 제2부회장' },
  { id: 6, name: 'LG전자 시네빔 4K 쇼츠+스탠드', count: 1, drawer: '이윤아 제1부회장' },
  { id: 7, name: '[깜짝] 올리브영 2만원권', count: 100, drawer: '김미형 회장', noRedraw: true },
  { id: 8, name: '삼성전자 갤럭시Z폴드8(256GB)', count: 1, drawer: '김미형 회장' },
]

const STORAGE_KEY = 'kaorn-draw-2026-v2'

// ===== HELPERS =====
function shuffle<T>(arr: T[]): T[] {
  const a = [...arr]
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[a[i], a[j]] = [a[j], a[i]]
  }
  return a
}

function last4(phone: string): string {
  return phone.replace(/\D/g, '').slice(-4)
}

function parseExcel(data: Uint8Array): Participant[] {
  const wb = XLSX.read(data, { type: 'array' })
  const sheet = wb.Sheets[wb.SheetNames[0]]
  const rows: string[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 })
  return rows
    .slice(1)
    .filter((r) => r[1] && String(r[1]).trim())
    .map((r) => {
      const name = String(r[1] ?? '').trim()
      const phone = String(r[2] ?? '').trim()
      const p4 = last4(phone)
      return {
        key: `${name}_${p4}`,
        name,
        display: p4 ? `${name} (${p4})` : name,
      }
    })
}

// ===== AUDIO =====
let sfxAudio: HTMLAudioElement | null = null      // 두구두구 (카운트다운)
let fanfareAudio: HTMLAudioElement | null = null  // 팡파레 (당첨 발표)

function preloadAudio() {
  if (typeof window === 'undefined') return
  if (!sfxAudio) {
    sfxAudio = new Audio('/두구두구두구 당첨 효과음.mp3')
    sfxAudio.preload = 'auto'
  }
  if (!fanfareAudio) {
    fanfareAudio = new Audio('/fanfare.mp3')
    fanfareAudio.preload = 'auto'
  }
}

let sfxStopTimer: ReturnType<typeof setTimeout> | null = null

function playCountdownAudio() {
  preloadAudio()
  if (!sfxAudio) return
  sfxAudio.currentTime = 1  // 1초부터 재생 시 5초 카운트다운과 타이밍 일치
  sfxAudio.volume = 0.9
  sfxAudio.play().catch(() => {})
  // 5초 후 자동 종료 (카운트다운 길이와 일치)
  if (sfxStopTimer) clearTimeout(sfxStopTimer)
  sfxStopTimer = setTimeout(() => { stopCountdownAudio() }, 5000)
}

function stopCountdownAudio() {
  if (!sfxAudio) return
  sfxAudio.pause()
  sfxAudio.currentTime = 0
}

function playFanfare() {
  preloadAudio()
  if (!fanfareAudio) return
  fanfareAudio.currentTime = 0
  fanfareAudio.volume = 0.9
  fanfareAudio.play().catch(() => {})
}

// 하위 호환
function playAudio() { playFanfare() }
function stopAudioAt(_sec: number) { /* 미사용 */ }

// ===== CONFETTI =====
function fireConfetti(ms = 2500) {
  const canvas = document.getElementById('confetti-canvas') as HTMLCanvasElement | null
  if (!canvas) return
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  canvas.width = window.innerWidth
  canvas.height = window.innerHeight
  const colors = ['#f59e0b', '#fbbf24', '#ef4444', '#10b981', '#8b5cf6', '#ec4899', '#06b6d4', '#ffffff', '#a78bfa']
  const particles: {
    x: number; y: number; vx: number; vy: number
    size: number; color: string; rotation: number; vr: number; shape: number
  }[] = []
  for (let i = 0; i < 280; i++) {
    particles.push({
      x: window.innerWidth * (0.1 + Math.random() * 0.8),
      y: window.innerHeight * (0.05 + Math.random() * 0.4),
      vx: (Math.random() - 0.5) * 22,
      vy: -Math.random() * 24 - 5,
      size: Math.random() * 14 + 4,
      color: colors[Math.floor(Math.random() * colors.length)],
      rotation: Math.random() * 360,
      vr: (Math.random() - 0.5) * 15,
      shape: Math.floor(Math.random() * 3),
    })
  }
  const fps = 60
  const maxFrames = Math.floor((ms / 1000) * fps)
  let frame = 0
  function animate() {
    if (frame >= maxFrames) {
      ctx!.clearRect(0, 0, canvas!.width, canvas!.height)
      return
    }
    ctx!.clearRect(0, 0, canvas!.width, canvas!.height)
    const life = Math.max(0, 1 - frame / maxFrames)
    particles.forEach((p) => {
      p.x += p.vx; p.vy += 0.35; p.y += p.vy; p.vx *= 0.985; p.rotation += p.vr
      ctx!.save()
      ctx!.translate(p.x, p.y)
      ctx!.rotate((p.rotation * Math.PI) / 180)
      ctx!.globalAlpha = life
      ctx!.fillStyle = p.color
      if (p.shape === 0) {
        ctx!.fillRect(-p.size / 2, -p.size / 2, p.size, p.size * 0.6)
      } else if (p.shape === 1) {
        ctx!.beginPath(); ctx!.arc(0, 0, p.size / 2, 0, Math.PI * 2); ctx!.fill()
      } else {
        ctx!.beginPath(); ctx!.moveTo(0, -p.size / 2); ctx!.lineTo(p.size / 2, p.size / 2)
        ctx!.lineTo(-p.size / 2, p.size / 2); ctx!.closePath(); ctx!.fill()
      }
      ctx!.restore()
    })
    frame++
    requestAnimationFrame(animate)
  }
  animate()
}

// ===== STATE PERSISTENCE =====
interface AppState {
  prizes: Prize[]
  confirmed: Record<number, Participant[]>
  participants: Participant[]
  fileName: string
}

function loadState(): Partial<AppState> {
  if (typeof window === 'undefined') return {}
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    return raw ? JSON.parse(raw) : {}
  } catch { return {} }
}

function saveState(s: AppState) {
  if (typeof window === 'undefined') return
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(s)) } catch {}
}

// ===== COUNTDOWN OVERLAY =====
function CountdownOverlay({ onComplete }: { onComplete: () => void }) {
  const [count, setCount] = useState(5)
  const onCompleteRef = useRef(onComplete)
  onCompleteRef.current = onComplete

  useEffect(() => {
    // 마운트 시 두구두구 효과음 재생
    playCountdownAudio()

    let alive = true
    let n = 5
    const tick = () => {
      if (!alive) return
      n--
      if (n <= 0) {
        setCount(0)
        stopCountdownAudio()   // 두구두구 정지
        playFanfare()          // 팡파레 재생
        setTimeout(() => { if (alive) onCompleteRef.current() }, 150)
        return
      }
      setCount(n)
      setTimeout(tick, 1000)
    }
    const t = setTimeout(tick, 1000)
    return () => {
      alive = false
      clearTimeout(t)
      // 컴포넌트 unmount 시에만 정지 (재추첨 등 취소 케이스)
      stopCountdownAudio()
    }
  }, [])

  const isWarm = count <= 2
  const numberColor = isWarm ? '#fbbf24' : '#ffffff'
  const glow = isWarm
    ? '0 0 40px rgba(251,191,36,1), 0 0 80px rgba(251,191,36,0.6), 0 0 180px rgba(251,191,36,0.3)'
    : '0 0 40px rgba(255,255,255,0.8), 0 0 80px rgba(255,255,255,0.4)'

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center select-none"
      style={{ background: 'rgba(5, 10, 40, 0.94)', backdropFilter: 'blur(10px)' }}
    >
      {count > 0 && (
        <div className="relative flex items-center justify-center">
          {/* Pulse ring */}
          <div
            key={`ring-${count}`}
            className="countdown-ring absolute rounded-full"
            style={{
              width: 'clamp(220px, 32vw, 360px)',
              height: 'clamp(220px, 32vw, 360px)',
              border: `4px solid ${numberColor}`,
              opacity: 0.5,
            }}
          />
          {/* Number */}
          <div
            key={count}
            className="countdown-number font-black"
            style={{
              fontSize: 'clamp(160px, 24vw, 280px)',
              lineHeight: 1,
              color: numberColor,
              textShadow: glow,
            }}
          >
            {count}
          </div>
        </div>
      )}
    </div>
  )
}

// ===== MAIN COMPONENT =====
export default function Home() {
  const [loaded, setLoaded] = useState(false)
  const [mode, setMode] = useState<Mode>('setup')
  const [drawPage, setDrawPage] = useState<DrawPage>('home')
  const [activePrizeId, setActivePrizeId] = useState<number>(1)

  // Config state
  const [prizes, setPrizes] = useState<Prize[]>(DEFAULT_PRIZES)
  const [participants, setParticipants] = useState<Participant[]>([])
  const [fileName, setFileName] = useState('')
  const [confirmed, setConfirmed] = useState<Record<number, Participant[]>>({})

  // Draw state
  const [drawPhase, setDrawPhase] = useState<DrawPhase>('idle')
  const [drumName, setDrumName] = useState('')
  const [pending, setPending] = useState<Participant[]>([])
  const [pendingConfirmed, setPendingConfirmed] = useState<Set<string>>(new Set())
  const [redrawingIdx, setRedrawingIdx] = useState<number | null>(null)
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null)
  // 재추첨 모드: 'single'(단일), 'batch'(미확인 전체), null(최초추첨)
  const redrawModeRef = useRef<'single' | 'batch' | null>(null)
  const redrawIdxRef = useRef<number>(-1)

  // Results
  const [copied, setCopied] = useState(false)

  // Load saved state
  useEffect(() => {
    const saved = loadState()
    if (saved.prizes) setPrizes(saved.prizes)
    if (saved.confirmed) setConfirmed(saved.confirmed)
    if (saved.participants) setParticipants(saved.participants)
    if (saved.fileName) setFileName(saved.fileName)
    setLoaded(true)
  }, [])

  // Save state on changes
  useEffect(() => {
    if (!loaded) return
    saveState({ prizes, confirmed, participants, fileName })
  }, [loaded, prizes, confirmed, participants, fileName])

  // All confirmed keys across all prizes
  const allConfirmedKeys = useCallback(() => {
    return new Set(Object.values(confirmed).flat().map((p) => p.key))
  }, [confirmed])

  // Pool for a given prize: excludes ALL previously confirmed winners
  const getPool = useCallback((prizeId: number) => {
    void prizeId
    const excludedKeys = allConfirmedKeys()
    return participants.filter((p) => !excludedKeys.has(p.key))
  }, [participants, allConfirmedKeys])

  const activePrize = prizes.find((p) => p.id === activePrizeId)
  const activePrizeConfirmed = confirmed[activePrizeId] ?? []
  const remaining = (activePrize?.count ?? 0) - activePrizeConfirmed.length

  const isPrizeDone = useCallback((prizeId: number) => {
    const prize = prizes.find((p) => p.id === prizeId)
    if (!prize) return false
    return (confirmed[prizeId]?.length ?? 0) >= prize.count
  }, [prizes, confirmed])

  // ---- DRAW LOGIC ----
  const startDraw = () => {
    const pool = getPool(activePrizeId)
    if (pool.length === 0 || remaining <= 0) return
    setPending([])
    setPendingConfirmed(new Set())
    setDrawPhase('countdown')
  }

  const onCountdownComplete = useCallback(() => {
    const mode = redrawModeRef.current
    redrawModeRef.current = null  // 모드 초기화

    if (mode === 'single') {
      // 단일 재추첨
      const idx = redrawIdxRef.current
      const excludeKeys = new Set([
        ...Object.values(confirmed).flat().map((p) => p.key),
        ...pending.filter((_, i) => i !== idx).map((p) => p.key),
      ])
      const pool = participants.filter((p) => !excludeKeys.has(p.key))
      const winner = shuffle(pool)[0]
      if (winner) {
        setPending((prev) => prev.map((p, i) => (i === idx ? winner : p)))
      }
      setRedrawingIdx(null)
      setDrawPhase('reveal')
      setTimeout(() => fireConfetti(2500), 200)
      return
    }

    if (mode === 'batch') {
      // 미확인자 일괄 재추첨
      const keepKeys = new Set(
        pending.filter((p) => pendingConfirmed.has(p.key)).map((p) => p.key)
      )
      const excludeKeys = new Set([
        ...Object.values(confirmed).flat().map((p) => p.key),
        ...keepKeys,
      ])
      const pool = participants.filter((p) => !excludeKeys.has(p.key))
      const unconfirmedCount = pending.filter((p) => !pendingConfirmed.has(p.key)).length
      const newWinners = shuffle(pool).slice(0, Math.min(unconfirmedCount, pool.length))
      let ni = 0
      setPending((prev) =>
        prev.map((p) =>
          !pendingConfirmed.has(p.key) && ni < newWinners.length
            ? newWinners[ni++]
            : p
        )
      )
      setRedrawingIdx(null)
      setDrawPhase('reveal')
      setTimeout(() => fireConfetti(2500), 200)
      return
    }

    // 최초 추첨
    const pool = getPool(activePrizeId)
    const drawCount = Math.min(remaining, pool.length)
    if (drawCount <= 0) { setDrawPhase('idle'); return }
    const winners = shuffle(pool).slice(0, drawCount)
    setPending(winners)
    setDrawPhase('reveal')
    setTimeout(() => fireConfetti(2500), 200)
  }, [activePrizeId, getPool, remaining, confirmed, pending, pendingConfirmed, participants])

  const confirmWinner = async (p: Participant) => {
    setConfirmed((prev) => ({
      ...prev,
      [activePrizeId]: [...(prev[activePrizeId] ?? []), p],
    }))
    setPendingConfirmed((prev) => new Set(prev).add(p.key))
    if (!activePrize) return
    await getSupabase().from('winners').insert({
      name: p.display,
      prize: activePrize.name,
      prize_type: 'lucky',
    })
  }

  const undoConfirm = async (p: Participant) => {
    setConfirmed((prev) => ({
      ...prev,
      [activePrizeId]: (prev[activePrizeId] ?? []).filter((c) => c.key !== p.key),
    }))
    setPendingConfirmed((prev) => { const n = new Set(prev); n.delete(p.key); return n })
    if (!activePrize) return
    await getSupabase().from('winners').delete()
      .eq('name', p.display).eq('prize', activePrize.name)
  }

  const startSingleRedraw = (idx: number) => {
    const excludeKeys = new Set([
      ...allConfirmedKeys(),
      ...pending.map((p) => p.key),
    ])
    const pool = participants.filter((p) => !excludeKeys.has(p.key))
    if (pool.length === 0) return
    redrawModeRef.current = 'single'
    redrawIdxRef.current = idx
    setDrawPhase('countdown')
  }

  const redrawUnconfirmed = () => {
    const unconfirmed = pending.filter((p) => !pendingConfirmed.has(p.key))
    if (unconfirmed.length === 0) return
    redrawModeRef.current = 'batch'
    setDrawPhase('countdown')
  }

  const allPendingConfirmed =
    pending.length > 0 && pending.every((p) => pendingConfirmed.has(p.key))

  const finishBatch = () => {
    setPending([])
    setPendingConfirmed(new Set())
    setDrawPhase('idle')
    // 추첨 완료 시 다음 미완료 경품으로 직접 이동, 없으면 목록으로
    if (isPrizeDone(activePrizeId)) {
      const nextPrize = prizes.find(
        (p) => p.id > activePrizeId && (confirmed[p.id]?.length ?? 0) < p.count
      )
      if (nextPrize) {
        setActivePrizeId(nextPrize.id)
        // drawPage는 'prize' 유지 — 다음 경품 화면으로 바로 전환
      } else {
        setDrawPage('home')
      }
    }
  }

  const handleReset = () => {
    if (!confirm('모든 데이터를 초기화하시겠습니까?\n(명단, 당첨 기록 모두 삭제됩니다)')) return
    localStorage.removeItem(STORAGE_KEY)
    setPrizes(DEFAULT_PRIZES)
    setConfirmed({})
    setParticipants([])
    setFileName('')
    setPending([])
    setPendingConfirmed(new Set())
    setDrawPhase('idle')
    setMode('setup')
    setDrawPage('home')
    setActivePrizeId(1)
  }

  const copyResults = () => {
    let text = '=== 2026 KAORN 경품 추첨 당첨자 ===\n\n'
    prizes.forEach((prize) => {
      const winners = confirmed[prize.id] ?? []
      text += `[${prize.id}순위] ${prize.name}\n`
      if (winners.length === 0) {
        text += '  미추첨\n'
      } else {
        winners.forEach((w, i) => { text += `  ${i + 1}. ${w.display}\n` })
      }
      text += '\n'
    })
    navigator.clipboard.writeText(text)
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }

  useEffect(() => {
    if (mode === 'draw') preloadAudio()
  }, [mode])

  if (!loaded) return (
    <div className="min-h-screen flex items-center justify-center" style={{ background: '#0a1628' }}>
      <p className="text-white/40">로딩 중...</p>
    </div>
  )

  // ==================== SETUP MODE ====================
  if (mode === 'setup') {
    return (
      <div className="min-h-screen" style={{ background: '#0a1628' }}>
        <header className="py-6 px-4 border-b border-white/10">
          <h1 className="text-center text-xl font-bold text-white">추첨 관리자 설정</h1>
          <p className="text-center text-sm mt-1 text-white/40">명단 업로드 후 추첨 화면으로 전환하세요</p>
        </header>

        <main className="max-w-2xl mx-auto p-6 space-y-5">
          {/* File upload */}
          <div
            className="rounded-2xl p-6"
            style={{ background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.1)' }}
          >
            <h2 className="text-base font-bold text-white mb-1">참가자 명단</h2>
            <p className="text-sm text-white/40 mb-4">엑셀 형식: B열(이름), C열(연락처)</p>
            <div className="flex items-center gap-3">
              <label
                className="cursor-pointer text-sm font-medium text-white px-4 py-2 rounded-lg transition-all hover:opacity-80"
                style={{ background: 'rgba(59,130,246,0.3)', border: '1px solid rgba(59,130,246,0.5)' }}
              >
                파일 선택
                <input
                  type="file" accept=".xlsx,.xls" className="hidden"
                  onChange={(e) => {
                    const file = e.target.files?.[0]
                    if (!file) return
                    setFileName(file.name)
                    const reader = new FileReader()
                    reader.onload = (ev) =>
                      setParticipants(parseExcel(new Uint8Array(ev.target?.result as ArrayBuffer)))
                    reader.readAsArrayBuffer(file)
                  }}
                />
              </label>
              {fileName && <span className="text-sm text-white/60">{fileName}</span>}
            </div>
            {participants.length > 0 && (
              <div
                className="mt-4 p-3 rounded-lg"
                style={{ background: 'rgba(16,185,129,0.12)', border: '1px solid rgba(16,185,129,0.3)' }}
              >
                <span className="text-sm font-medium text-emerald-400">{participants.length}명 로드 완료</span>
                <div className="mt-2 max-h-36 overflow-y-auto text-xs text-white/50 space-y-1">
                  {participants.map((p, i) => (
                    <div key={i}>{i + 1}. {p.display}</div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Prize config */}
          <div
            className="rounded-2xl p-6"
            style={{ background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.1)' }}
          >
            <h2 className="text-base font-bold text-white mb-4">경품 목록 설정</h2>
            <div className="space-y-2">
              {prizes.map((prize, idx) => (
                <div
                  key={prize.id}
                  className="flex items-center gap-2 p-3 rounded-xl"
                  style={{ background: 'rgba(255,255,255,0.04)' }}
                >
                  <span className="text-white/30 text-xs w-5 shrink-0 text-center">{idx + 1}</span>
                  <span className="text-white text-sm flex-1 min-w-0 truncate">{prize.name}</span>
                  <div className="flex items-center gap-1 shrink-0">
                    <span className="text-white/30 text-xs">수량</span>
                    <input
                      type="number" min={1} max={20} value={prize.count}
                      onChange={(e) =>
                        setPrizes((prev) =>
                          prev.map((p) => p.id === prize.id ? { ...p, count: Number(e.target.value) } : p)
                        )
                      }
                      className="w-14 text-center text-sm text-white rounded-lg px-2 py-1 outline-none"
                      style={{ background: 'rgba(255,255,255,0.08)', border: '1px solid rgba(255,255,255,0.15)' }}
                    />
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <span className="text-white/30 text-xs">추첨자</span>
                    <input
                      type="text" value={prize.drawer}
                      onChange={(e) =>
                        setPrizes((prev) =>
                          prev.map((p) => p.id === prize.id ? { ...p, drawer: e.target.value } : p)
                        )
                      }
                      className="w-36 text-sm text-white rounded-lg px-2 py-1 outline-none"
                      style={{ background: 'rgba(255,255,255,0.08)', border: '1px solid rgba(255,255,255,0.15)' }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>

          <button
            onClick={() => { setMode('draw'); setDrawPage('home') }}
            disabled={participants.length === 0}
            className="w-full text-white text-lg font-bold py-4 rounded-2xl transition-all disabled:opacity-30 hover:opacity-90"
            style={{ background: 'linear-gradient(135deg, #1a4fa0, #2563eb)' }}
          >
            {participants.length > 0 ? '추첨 화면으로 전환' : '명단을 먼저 업로드하세요'}
          </button>

          {(participants.length > 0 || Object.keys(confirmed).length > 0) && (
            <button
              onClick={handleReset}
              className="w-full text-white font-bold py-3 rounded-2xl transition-all hover:opacity-80"
              style={{ background: 'rgba(239,68,68,0.2)', border: '1px solid rgba(239,68,68,0.4)' }}
            >
              전체 초기화
            </button>
          )}
        </main>
      </div>
    )
  }

  // ==================== DRAW MODE ====================
  const bg = 'linear-gradient(180deg, #060e22 0%, #0a1628 40%, #0d1b3e 70%, #0a1628 100%)'

  return (
    <div className="min-h-screen" style={{ background: bg }}>
      {/* Countdown overlay */}
      {drawPhase === 'countdown' && (
        <CountdownOverlay onComplete={onCountdownComplete} />
      )}

      {/* Header */}
      <header
        className="py-4 px-4 relative"
        style={{ borderBottom: '1px solid rgba(255,255,255,0.08)' }}
      >
        <div className="text-center">
          <div className="text-white text-xs font-medium tracking-widest uppercase mb-0.5">
            2026 KAORN 병원수술간호사회
          </div>
          <h1 className={`text-white font-bold ${drawPage === 'prize' ? 'text-xl md:text-2xl' : 'text-base md:text-lg'}`}>
            제37회 학술대회 및 제12회 Surgical EXPO <span style={{ color: '#fbbf24' }}>경품추첨</span>
          </h1>
        </div>
        <button
          onClick={() => setMode('setup')}
          className="absolute top-4 right-4 text-xs text-white/30 hover:text-white/70 px-3 py-1.5 rounded-lg transition-colors"
          style={{ background: 'rgba(255,255,255,0.06)' }}
        >
          관리자
        </button>
      </header>

      {/* noRedraw 100명 확정 버튼 — 화면 하단 고정 */}
      {drawPage === 'prize' && activePrize?.noRedraw && drawPhase === 'reveal' && pending.length > 0 && (
        <div
          className="fixed bottom-0 left-0 right-0 flex justify-center py-4 z-40"
          style={{ background: 'linear-gradient(to top, rgba(6,14,34,0.98) 70%, transparent)' }}
        >
          <button
            onClick={async () => {
              for (const p of pending) { await confirmWinner(p) }
              finishBatch()
            }}
            className="text-white font-black py-4 px-16 rounded-full transition-all hover:scale-105"
            style={{
              fontSize: 'clamp(1rem, 1.5vw, 1.3rem)',
              background: 'rgba(99,102,241,0.5)',
              border: '2px solid rgba(139,92,246,0.7)',
              boxShadow: '0 0 40px rgba(99,102,241,0.5)',
            }}
          >
            전체 확정 ({pending.length}명) → 다음 추첨으로
          </button>
        </div>
      )}

      <main className={drawPage === 'prize' ? 'w-full px-6 md:px-12 py-4 md:py-6' : 'max-w-4xl mx-auto p-4 md:p-6'}>

        {/* ===== HOME PAGE ===== */}
        {drawPage === 'home' && (
          <div className="space-y-3 mt-4">
            <p className="text-white/30 text-sm text-center mb-5">추첨할 경품을 선택하세요</p>
            {prizes.map((prize, idx) => {
              const done = isPrizeDone(prize.id)
              const confirmedCount = confirmed[prize.id]?.length ?? 0
              return (
                <button
                  key={prize.id}
                  onClick={() => {
                    setActivePrizeId(prize.id)
                    setDrawPage('prize')
                    setDrawPhase('idle')
                    setPending([])
                    setPendingConfirmed(new Set())
                  }}
                  className="w-full text-left p-5 rounded-2xl transition-all group hover:scale-[1.01]"
                  style={{
                    background: done
                      ? 'rgba(16,185,129,0.08)'
                      : 'rgba(255,255,255,0.04)',
                    border: done
                      ? '1px solid rgba(16,185,129,0.3)'
                      : '1px solid rgba(255,255,255,0.08)',
                  }}
                >
                  <div className="flex items-center gap-4">
                    <span
                      className="text-2xl md:text-3xl font-black shrink-0 tabular-nums"
                      style={{ color: done ? 'rgba(52,211,153,0.6)' : 'rgba(255,255,255,0.15)' }}
                    >
                      {String(idx + 1).padStart(2, '0')}
                    </span>
                    <div className="flex-1 min-w-0">
                      <div
                        className="font-bold text-base md:text-lg leading-tight"
                        style={{ color: done ? '#6ee7b7' : 'white' }}
                      >
                        {prize.name}
                      </div>
                      <div className="text-sm mt-1 flex items-center gap-3">
                        <span className="text-white/30">추첨: {prize.drawer}</span>
                        <span style={{ color: done ? '#6ee7b7' : 'rgba(255,255,255,0.3)' }}>
                          {confirmedCount}/{prize.count}명
                        </span>
                      </div>
                    </div>
                    {done
                      ? <span className="text-emerald-400 text-sm font-bold shrink-0">완료 ✓</span>
                      : <span className="text-white/20 group-hover:text-white/50 transition-colors shrink-0 text-lg">›</span>
                    }
                  </div>
                </button>
              )
            })}
            <button
              onClick={() => setDrawPage('results')}
              className="w-full p-4 rounded-2xl text-white/30 hover:text-white/60 transition-all text-sm mt-2"
              style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.06)' }}
            >
              최종 결과 보기
            </button>

            {Object.keys(confirmed).length > 0 && (
              <button
                onClick={() => {
                  if (!confirm('추첨 내역을 초기화하시겠습니까?\n(참가자 명단은 유지됩니다)')) return
                  setConfirmed({})
                  setPending([])
                  setPendingConfirmed(new Set())
                  setDrawPhase('idle')
                }}
                className="w-full p-3 rounded-2xl text-sm transition-all hover:opacity-80 mt-1"
                style={{
                  background: 'rgba(239,68,68,0.08)',
                  border: '1px solid rgba(239,68,68,0.2)',
                  color: 'rgba(239,68,68,0.6)',
                }}
              >
                추첨 내역 초기화
              </button>
            )}
          </div>
        )}

        {/* ===== PRIZE DRAW PAGE ===== */}
        {drawPage === 'prize' && activePrize && (
          <div className={`space-y-4 flex flex-col ${activePrize.noRedraw && drawPhase === 'reveal' ? '' : 'min-h-[80vh]'}`}>
            <button
              onClick={() => {
                setDrawPage('home')
                setDrawPhase('idle')
                setPending([])
                setPendingConfirmed(new Set())
              }}
              className="text-sm text-white/30 hover:text-white/60 transition-colors self-start"
            >
              ← 목록으로
            </button>

            {/* Prize info */}
            <div
              className="text-center rounded-3xl"
              style={{
                background: 'rgba(255,255,255,0.03)',
                border: '1px solid rgba(255,255,255,0.08)',
                padding: activePrize?.noRedraw && drawPhase === 'reveal' ? '1.2rem 2rem' : 'clamp(1.5rem, 4vw, 2.5rem) 2rem',
              }}
            >
              {/* noRedraw reveal 시 한 줄 압축 표시 */}
              {activePrize.noRedraw && drawPhase === 'reveal' ? (
                <div className="flex items-center justify-center gap-4 flex-wrap">
                  <span
                    className="font-black rounded-full px-3 py-1"
                    style={{ fontSize: '0.95rem', background: 'rgba(99,102,241,0.25)', color: '#c4b5fd', border: '1.5px solid rgba(139,92,246,0.5)' }}
                  >
                    {activePrize.id}번 경품
                  </span>
                  <span className="text-white font-black" style={{ fontSize: 'clamp(1rem, 2vw, 1.5rem)', wordBreak: 'keep-all' }}>
                    {activePrize.name}
                  </span>
                  <span style={{ fontSize: '1rem', color: '#93c5fd', fontWeight: 700 }}>
                    추첨자: {activePrize.drawer}
                  </span>
                </div>
              ) : (
                <>
                  <div
                    className="inline-block font-black rounded-full mb-6"
                    style={{
                      fontSize: 'clamp(1.1rem, 2vw, 1.6rem)',
                      padding: 'clamp(0.5rem, 1vw, 0.8rem) clamp(1.5rem, 3vw, 2.5rem)',
                      background: 'rgba(99,102,241,0.25)',
                      color: '#c4b5fd',
                      border: '2px solid rgba(139,92,246,0.5)',
                      boxShadow: '0 0 24px rgba(99,102,241,0.3)',
                    }}
                  >
                    {activePrize.id}번 경품
                  </div>
                  <h2
                    className="text-white font-black leading-snug mb-6"
                    style={{ fontSize: 'clamp(1.8rem, 3.8vw, 3.8rem)', wordBreak: 'keep-all', overflowWrap: 'anywhere' }}
                  >
                    {activePrize.name}
                  </h2>
                  <p style={{ fontSize: 'clamp(1.3rem, 2.4vw, 2rem)', color: '#93c5fd', fontWeight: 700 }}>
                    추첨자: {activePrize.drawer}
                  </p>
                  <p className="mt-3" style={{ fontSize: 'clamp(0.9rem, 1.5vw, 1.2rem)', color: 'rgba(255,255,255,0.35)' }}>
                    {activePrizeConfirmed.length}/{activePrize.count}명 확정
                    {remaining > 0 && ` · 추첨 대상 ${remaining}명`}
                  </p>
                </>
              )}
            </div>

            {/* DRUMROLL phase */}
            {drawPhase === 'drumroll' && (
              <div className="flex flex-col items-center justify-center flex-1 py-16">
                <div
                  className="font-black text-white drumroll-active"
                  style={{
                    fontSize: 'clamp(3rem, 10vw, 8rem)',
                    textShadow: '0 0 50px rgba(251,191,36,0.6), 0 0 100px rgba(99,102,241,0.3)',
                  }}
                >
                  {drumName || '...'}
                </div>
                <p className="mt-8 text-white/30 tracking-widest text-lg uppercase">추첨 중...</p>
              </div>
            )}

            {/* IDLE phase: start button */}
            {drawPhase === 'idle' && pending.length === 0 && (
              <div className="flex flex-col items-center flex-1 justify-center gap-6 py-10">
                {remaining > 0 ? (
                  <button
                    onClick={startDraw}
                    disabled={getPool(activePrizeId).length === 0}
                    className="text-white font-black rounded-full transition-all disabled:opacity-30 hover:scale-105 active:scale-95"
                    style={{
                      fontSize: 'clamp(1.5rem, 3vw, 2.5rem)',
                      padding: 'clamp(1.5rem, 3vw, 2.2rem) clamp(4rem, 10vw, 10rem)',
                      background: 'linear-gradient(135deg, #1a3a8a, #4c1d95, #1a3a8a)',
                      boxShadow: '0 0 60px rgba(99,102,241,0.5), 0 0 120px rgba(99,102,241,0.2)',
                    }}
                  >
                    추첨 시작
                  </button>
                ) : (
                  <div className="text-emerald-400 font-bold py-10" style={{ fontSize: 'clamp(1.5rem, 3vw, 2.5rem)' }}>
                    🎉 추첨 완료!
                  </div>
                )}
                {getPool(activePrizeId).length === 0 && remaining > 0 && (
                  <p className="text-white/30 text-lg">추첨 가능한 인원이 없습니다</p>
                )}
              </div>
            )}

            {/* REVEAL phase: winner cards */}
            {drawPhase === 'reveal' && pending.length > 0 && (
              <div>
                {/* noRedraw 모드에서는 헤딩을 작게 */}
                <h3
                  className="text-center font-black text-white"
                  style={{
                    fontSize: activePrize?.noRedraw ? 'clamp(1.2rem, 2vw, 1.8rem)' : 'clamp(2rem, 4vw, 3.5rem)',
                    marginBottom: activePrize?.noRedraw ? '0.5rem' : '2rem',
                  }}
                >
                  🎉 당첨을 축하합니다!
                </h3>

                {/* ── noRedraw 모드: 콤팩트 그리드 (올리브영 100명 등) ── */}
                {activePrize?.noRedraw ? (
                  <div>
                    <p className="text-center text-white/40 text-xs mb-2">총 {pending.length}명 당첨</p>
                    {/* 14열 × 행높이 균등 분배 — 뷰포트 꽉 채움 (하단 버튼 바 여유) */}
                    <div
                      className="grid"
                      style={{
                        gridTemplateColumns: 'repeat(14, 1fr)',
                        gridAutoRows: '1fr',
                        gap: '5px',
                        height: 'calc(100vh - 330px)',
                      }}
                    >
                      {pending.map((p, i) => {
                        const namePart = p.display.includes(' (') ? p.display.split(' (')[0] : p.display
                        const phonePart = p.display.includes(' (') ? '(' + p.display.split(' (')[1] : ''
                        return (
                        <div
                          key={i}
                          className="rounded-md flex flex-col items-center justify-center winner-reveal"
                          style={{
                            background: 'rgba(99,102,241,0.18)',
                            border: '1px solid rgba(139,92,246,0.35)',
                            animationDelay: `${Math.min(i * 0.02, 1.5)}s`,
                          }}
                        >
                          <span className="text-white font-black text-center leading-tight" style={{ fontSize: 'clamp(0.75rem, 1.1vw, 1rem)' }}>{namePart}</span>
                          {phonePart && <span className="text-white/70 font-bold text-center leading-tight" style={{ fontSize: 'clamp(0.7rem, 1vw, 0.9rem)' }}>{phonePart}</span>}
                        </div>
                        )
                      })}
                    </div>
                  </div>
                ) : (
                  /* ── 일반 모드: 현장확인 / 재추첨 카드 ── */
                  <div>
                    <div
                      className={`grid gap-5 mb-8 ${
                        pending.length === 1 ? 'grid-cols-1 max-w-md mx-auto' : 'grid-cols-1 md:grid-cols-2'
                      }`}
                    >
                      {pending.map((p, i) => {
                        const isConfirmed = pendingConfirmed.has(p.key)
                        const isTmp = p.key.startsWith('_tmp_')
                        const isRedrawing =
                          (redrawingIdx === i || (redrawingIdx === -1 && !isConfirmed && isTmp))

                        return (
                          <div
                            key={i}
                            className={`rounded-2xl p-7 text-center ${
                              isRedrawing ? 'card-drumroll' : isConfirmed ? '' : 'winner-reveal'
                            }`}
                            style={{
                              background: isRedrawing
                                ? 'rgba(99,102,241,0.2)'
                                : isConfirmed
                                ? 'rgba(16,185,129,0.15)'
                                : 'linear-gradient(135deg, rgba(99,102,241,0.3), rgba(124,58,237,0.3))',
                              border: isRedrawing
                                ? '1px solid rgba(99,102,241,0.4)'
                                : isConfirmed
                                ? '1px solid rgba(16,185,129,0.4)'
                                : '1px solid rgba(139,92,246,0.4)',
                              animationDelay: `${i * 0.15}s`,
                              boxShadow: isConfirmed ? 'none' : '0 0 40px rgba(99,102,241,0.15)',
                            }}
                          >
                            <div
                              className="font-black mb-5"
                              style={{ fontSize: 'clamp(1.8rem, 4vw, 3.5rem)', color: isConfirmed ? '#6ee7b7' : 'white' }}
                            >
                              {isRedrawing ? p.name : p.display}
                            </div>

                            {isConfirmed ? (
                              <div className="flex items-center justify-center gap-3">
                                <span className="text-emerald-400 text-sm font-bold">현장 확인 완료 ✓</span>
                                <button
                                  onClick={() => undoConfirm(p)}
                                  className="text-xs text-white/30 hover:text-white/60 px-3 py-1 rounded-full transition-colors"
                                  style={{ background: 'rgba(255,255,255,0.08)' }}
                                >
                                  취소
                                </button>
                              </div>
                            ) : isRedrawing ? (
                              <p className="text-white/30 text-sm">재추첨 중...</p>
                            ) : (
                              <div className="flex gap-3 justify-center">
                                <button
                                  onClick={() => confirmWinner(p)}
                                  disabled={redrawingIdx !== null}
                                  className="font-bold text-white rounded-full disabled:opacity-30 transition-all hover:scale-105"
                                  style={{ fontSize: 'clamp(0.9rem, 1.5vw, 1.2rem)', padding: '0.75rem 2.5rem', background: 'rgba(16,185,129,0.4)', border: '1px solid rgba(16,185,129,0.6)' }}
                                >
                                  현장 확인
                                </button>
                                <button
                                  onClick={() => startSingleRedraw(i)}
                                  disabled={redrawingIdx !== null}
                                  className="font-bold text-white/60 rounded-full disabled:opacity-30 transition-all hover:scale-105"
                                  style={{ fontSize: 'clamp(0.9rem, 1.5vw, 1.2rem)', padding: '0.75rem 2.5rem', background: 'rgba(255,255,255,0.08)', border: '1px solid rgba(255,255,255,0.15)' }}
                                >
                                  재추첨
                                </button>
                              </div>
                            )}
                          </div>
                        )
                      })}
                    </div>

                    <div className="flex justify-center">
                      {allPendingConfirmed ? (
                        <button
                          onClick={finishBatch}
                          className="text-white font-black py-5 px-16 rounded-full transition-all hover:scale-105 text-lg"
                          style={{
                            background: 'linear-gradient(135deg, #1a3a8a, #4c1d95)',
                            boxShadow: '0 0 40px rgba(99,102,241,0.4)',
                          }}
                        >
                          {isPrizeDone(activePrizeId) ? '다음 추첨으로 →' : '계속 추첨'}
                        </button>
                      ) : (
                        <button
                          onClick={redrawUnconfirmed}
                          disabled={redrawingIdx !== null}
                          className="text-white font-bold py-4 px-10 rounded-full disabled:opacity-30 transition-all"
                          style={{ background: 'rgba(255,255,255,0.08)', border: '1px solid rgba(255,255,255,0.15)' }}
                        >
                          미확인자 재추첨 ({pending.filter((p) => !pendingConfirmed.has(p.key)).length}명)
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Confirmed winners for this prize */}
            {activePrizeConfirmed.length > 0 && (
              <div
                className="rounded-xl p-5"
                style={{ background: 'rgba(16,185,129,0.06)', border: '1px solid rgba(16,185,129,0.15)' }}
              >
                <h3 className="text-emerald-400 font-bold mb-3 text-sm">
                  확정 당첨자 ({activePrizeConfirmed.length}/{activePrize.count}명)
                </h3>
                <div className="space-y-2">
                  {activePrizeConfirmed.map((p, i) => (
                    <div
                      key={i}
                      className="text-white/70 text-sm py-2 px-3 rounded-lg"
                      style={{ background: 'rgba(255,255,255,0.04)' }}
                    >
                      {i + 1}. {p.display}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ===== RESULTS PAGE ===== */}
        {drawPage === 'results' && (
          <div className="space-y-5">
            <button
              onClick={() => setDrawPage('home')}
              className="text-sm text-white/30 hover:text-white/60 transition-colors"
            >
              ← 돌아가기
            </button>
            <div
              className="rounded-2xl p-6"
              style={{ background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.08)' }}
            >
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-white font-bold text-lg">전체 당첨자 목록</h2>
                <button
                  onClick={copyResults}
                  className="text-sm px-4 py-2 rounded-lg transition-colors font-medium"
                  style={{
                    background: copied ? 'rgba(16,185,129,0.25)' : 'rgba(255,255,255,0.08)',
                    color: copied ? '#6ee7b7' : 'white',
                    border: copied ? '1px solid rgba(16,185,129,0.4)' : '1px solid rgba(255,255,255,0.1)',
                  }}
                >
                  {copied ? '복사됨 ✓' : '결과 복사'}
                </button>
              </div>
              <div className="space-y-6">
                {prizes.map((prize) => {
                  const winners = confirmed[prize.id] ?? []
                  return (
                    <div key={prize.id}>
                      <div
                        className="flex items-center gap-2 mb-3 pb-2"
                        style={{ borderBottom: '1px solid rgba(255,255,255,0.08)' }}
                      >
                        <span className="text-white/30 text-sm font-bold">{prize.id}.</span>
                        <span className="text-white/70 text-sm font-bold">{prize.name}</span>
                        <span className="text-white/30 text-xs ml-auto">
                          {winners.length}/{prize.count}명
                        </span>
                      </div>
                      {winners.length === 0 ? (
                        <p className="text-white/20 text-sm px-1">미추첨</p>
                      ) : (
                        <div className="space-y-1">
                          {winners.map((w, i) => (
                            <div
                              key={i}
                              className="text-white/80 text-sm py-2 px-3 rounded-lg"
                              style={{ background: 'rgba(255,255,255,0.04)' }}
                            >
                              {i + 1}. {w.display}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            </div>
          </div>
        )}
      </main>

      <canvas id="confetti-canvas" />
    </div>
  )
}
