import { useState } from 'react'

export default function App() {
  const [book, setBook] = useState('시편')
  const [chapter, setChapter] = useState('23')
  const [isAllVerses, setIsAllVerses] = useState(true)
  const [startVerse, setStartVerse] = useState('1')
  const [endVerse, setEndVerse] = useState('')
  const [version, setVersion] = useState('개역개정')
  const [fontName, setFontName] = useState('경기천년제목 Bold')
  const [loading, setLoading] = useState(false)
  const [versesData, setVersesData] = useState([])
  const [activeSlideIndex, setActiveSlideIndex] = useState(0)

  const unit = book.trim() === '시편' || book.trim() === '시' ? '편' : '장'

  const handleSearch = async () => {
    setLoading(true)
    try {
      let url = `/api/search?book=${encodeURIComponent(book)}&chapter=${chapter}`
      if (!isAllVerses) {
        if (startVerse) url += `&start_verse=${startVerse}`
        if (endVerse) url += `&end_verse=${endVerse}`
      }

      const res = await fetch(url)
      const data = await res.json()
      if (!res.ok) {
        alert(data.detail || '본문 조회가 실패했습니다.')
        return
      }

      if (data.verses) {
        setVersesData(data.verses)
        setActiveSlideIndex(0)
      }
    } catch (err) {
      alert('백엔드 서버와 통신할 수 없습니다.')
    } finally {
      setLoading(false)
    }
  }

  const handleDownloadPPT = () => {
    let downloadUrl = `/api/generate-ppt?book=${encodeURIComponent(book)}&chapter=${chapter}&version=${encodeURIComponent(version)}&font_name=${encodeURIComponent(fontName)}`
    if (!isAllVerses) {
      if (startVerse) downloadUrl += `&start_verse=${startVerse}`
      if (endVerse) downloadUrl += `&end_verse=${endVerse}`
    }

    const link = document.createElement('a')
    link.href = downloadUrl
    link.setAttribute('download', `${book}_${chapter}${unit}.pptx`)
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  // 화면 프리뷰용 폰트 패밀리 결정
  const getPreviewFontFamily = () => {
    if (fontName.includes('경기천년바탕')) return "'GyeonggiBatang', serif"
    if (fontName.includes('경기천년제목V')) return "'GyeonggiTitleV', sans-serif"
    if (fontName.includes('경기천년제목')) return "'GyeonggiTitle', sans-serif"
    return "Pretendard, sans-serif"
  }

  const getPreviewFontWeight = () => {
    if (fontName.includes('Light')) return '300'
    if (fontName.includes('Regular')) return '400'
    if (fontName.includes('Medium')) return '500'
    return '700'
  }

  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#0f172a', color: '#f8fafc', padding: '32px 20px', fontFamily: 'Pretendard, sans-serif' }}>
      {/* 경기천년체 웹폰트 로드 */}
      <style>{`
        @import url('https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_one@1.0/GyeonggiBatang.woff');
        @import url('https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_one@1.0/GyeonggiTitle.woff');
        @import url('https://cdn.jsdelivr.net/gh/projectnoonnu/noonfonts_one@1.0/GyeonggiTitleV.woff');
      `}</style>

      <div style={{ maxWidth: '1100px', margin: '0 auto' }}>
        
        {/* 헤더 */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '28px', borderBottom: '1px solid #334155', paddingBottom: '20px' }}>
          <div>
            <h1 style={{ fontSize: '24px', fontWeight: '800', color: '#38bdf8', margin: '0 0 6px 0' }}>
              BibleScript Pro 📖
            </h1>
            <p style={{ fontSize: '14px', color: '#94a3b8', margin: 0 }}>
              예배용 고화질 슬라이드 자동 생성 (경기천년체 지원)
            </p>
          </div>
          <span style={{ backgroundColor: '#0369a1', color: '#e0f2fe', padding: '6px 14px', borderRadius: '20px', fontSize: '12px', fontWeight: 'bold' }}>
            Typography Ready
          </span>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '380px 1fr', gap: '28px' }}>
          
          {/* 컨트롤 패널 */}
          <div style={{ backgroundColor: '#1e293b', borderRadius: '16px', padding: '24px', border: '1px solid #334155' }}>
            <h2 style={{ fontSize: '17px', fontWeight: '700', marginBottom: '18px', color: '#e2e8f0' }}>본문 설정</h2>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', marginBottom: '24px' }}>
              <div>
                <label style={{ display: 'block', fontSize: '12px', color: '#94a3b8', marginBottom: '6px' }}>성경 (약어 가능)</label>
                <input 
                  type="text" 
                  value={book} 
                  onChange={(e) => setBook(e.target.value)} 
                  placeholder="창, 출, 시, 롬 등"
                  style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid #475569', backgroundColor: '#0f172a', color: 'white', boxSizing: 'border-box' }}
                />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '12px', color: '#94a3b8', marginBottom: '6px' }}>{unit} 번호</label>
                <input 
                  type="number" 
                  value={chapter} 
                  onChange={(e) => setChapter(e.target.value)} 
                  style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid #475569', backgroundColor: '#0f172a', color: 'white', boxSizing: 'border-box' }}
                />
              </div>

              <div style={{ backgroundColor: '#0f172a', padding: '10px 14px', borderRadius: '8px', border: '1px solid #334155' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 'bold', color: '#38bdf8', cursor: 'pointer' }}>
                  <input 
                    type="checkbox" 
                    checked={isAllVerses} 
                    onChange={(e) => setIsAllVerses(e.target.checked)} 
                    style={{ width: '16px', height: '16px', cursor: 'pointer' }}
                  />
                  ⚡ 이 {unit}의 전체 절 자동 생성
                </label>
              </div>

              {!isAllVerses && (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', color: '#94a3b8', marginBottom: '6px' }}>시작 절</label>
                    <input 
                      type="number" 
                      value={startVerse} 
                      onChange={(e) => setStartVerse(e.target.value)} 
                      style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid #475569', backgroundColor: '#0f172a', color: 'white', boxSizing: 'border-box' }}
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', color: '#94a3b8', marginBottom: '6px' }}>끝 절</label>
                    <input 
                      type="number" 
                      value={endVerse} 
                      onChange={(e) => setEndVerse(e.target.value)} 
                      style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid #475569', backgroundColor: '#0f172a', color: 'white', boxSizing: 'border-box' }}
                    />
                  </div>
                </div>
              )}

              {/* 폰트 선택 드롭다운 (경기천년체 6종 추가) */}
              <div>
                <label style={{ display: 'block', fontSize: '12px', color: '#94a3b8', marginBottom: '6px' }}>출력 폰트 선택</label>
                <select 
                  value={fontName} 
                  onChange={(e) => setFontName(e.target.value)}
                  style={{ width: '100%', padding: '10px', borderRadius: '8px', border: '1px solid #475569', backgroundColor: '#0f172a', color: 'white' }}
                >
                  <optgroup label="경기천년체">
                    <option value="경기천년바탕 Bold">경기천년바탕 Bold</option>
                    <option value="경기천년바탕 Regular">경기천년바탕 Regular</option>
                    <option value="경기천년제목 Bold">경기천년제목 Bold</option>
                    <option value="경기천년제목 Light">경기천년제목 Light</option>
                    <option value="경기천년제목 Medium">경기천년제목 Medium</option>
                    <option value="경기천년제목V Bold">경기천년제목V Bold</option>
                  </optgroup>
                  <optgroup label="기본 서체">
                    <option value="Pretendard">Pretendard</option>
                    <option value="맑은 고딕">맑은 고딕</option>
                  </optgroup>
                </select>
              </div>
            </div>

            <button 
              onClick={handleSearch}
              disabled={loading}
              style={{ width: '100%', padding: '14px', backgroundColor: '#0284c7', color: 'white', border: 'none', borderRadius: '10px', fontSize: '15px', fontWeight: 'bold', cursor: 'pointer', marginBottom: '10px' }}
            >
              {loading ? '본문 로딩 중...' : '1. 본문 확인 및 슬라이드 구성'}
            </button>

            <button 
              onClick={handleDownloadPPT}
              style={{ width: '100%', padding: '14px', backgroundColor: '#10b981', color: 'white', border: 'none', borderRadius: '10px', fontSize: '15px', fontWeight: 'bold', cursor: 'pointer' }}
            >
              2. 📊 완성 PPTX 파일 다운로드
            </button>
          </div>

          {/* 슬라이드 뷰어 */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h2 style={{ fontSize: '17px', fontWeight: '700', color: '#e2e8f0', margin: 0 }}>
                실제 슬라이드 뷰어 (16:9)
              </h2>
              {versesData.length > 0 && (
                <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                  <button 
                    onClick={() => setActiveSlideIndex(prev => Math.max(0, prev - 1))}
                    disabled={activeSlideIndex === 0}
                    style={{ padding: '6px 14px', borderRadius: '6px', backgroundColor: '#334155', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 'bold' }}
                  >
                    ◀ 이전
                  </button>
                  <span style={{ fontSize: '13px', color: '#94a3b8', minWidth: '70px', textAlign: 'center' }}>
                    {activeSlideIndex + 1} / {versesData.length}
                  </span>
                  <button 
                    onClick={() => setActiveSlideIndex(prev => Math.min(versesData.length - 1, prev + 1))}
                    disabled={activeSlideIndex === versesData.length - 1}
                    style={{ padding: '6px 14px', borderRadius: '6px', backgroundColor: '#334155', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 'bold' }}
                  >
                    다음 ▶
                  </button>
                </div>
              )}
            </div>

            <div style={{
              width: '100%',
              aspectRatio: '16 / 9',
              backgroundColor: '#002060',
              borderRadius: '16px',
              boxShadow: '0 10px 25px rgba(0,0,0,0.5)',
              padding: '40px 32px',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'flex-start',
              boxSizing: 'border-box'
            }}>
              {versesData.length > 0 ? (
                <>
                  <div style={{ 
                    fontFamily: getPreviewFontFamily(),
                    fontSize: '28px', 
                    fontWeight: '800', 
                    color: '#FFC000', 
                    marginBottom: '24px', 
                    textAlign: 'center' 
                  }}>
                    {book} {chapter}{unit} {versesData[activeSlideIndex].verse_num}절
                  </div>
                  <div style={{ 
                    fontFamily: getPreviewFontFamily(),
                    fontSize: '24px', 
                    fontWeight: getPreviewFontWeight(), 
                    color: '#FFFFFF', 
                    textAlign: 'center', 
                    lineHeight: '1.6', 
                    wordBreak: 'keep-all', 
                    maxWidth: '88%' 
                  }}>
                    {versesData[activeSlideIndex].text}
                  </div>
                </>
              ) : (
                <div style={{ height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#94a3b8', fontSize: '16px' }}>
                  성경과 장을 입력하고 [본문 확인 및 슬라이드 구성]을 누르면 여기에 슬라이드가 나타납니다.
                </div>
              )}
            </div>
          </div>

        </div>

      </div>
    </div>
  )
}