import React, { useState, useEffect } from "react";

export default function App() {
  const [versions, setVersions] = useState(["개역개정"]);
  const [selectedVersions, setSelectedVersions] = useState(["개역개정"]);
  const [book, setBook] = useState("창세기");
  const [chapter, setChapter] = useState(1);
  const [startVerse, setStartVerse] = useState("");
  const [endVerse, setEndVerse] = useState("");
  const [fontName, setFontName] = useState("경기천년제목 Medium");

  // 여러 본문 일괄 입력 상태
  const [multiQuery, setMultiQuery] = useState("");

  // 키워드 검색 상태
  const [searchWord, setSearchWord] = useState("");
  const [searchResults, setSearchResults] = useState([]);
  const [isSearching, setIsSearching] = useState(false);

  // 구절 인라인 수정 상태
  const [editVerse, setEditVerse] = useState(null);
  const [editText, setEditText] = useState("");

  useEffect(() => {
    fetch("/api/versions")
      .then((r) => r.json())
      .then((data) => {
        if (data && Array.isArray(data.versions) && data.versions.length > 0) {
          setVersions(data.versions);
        }
      })
      .catch(console.error);
  }, []);

  const toggleVersion = (ver) => {
    if (selectedVersions.includes(ver)) {
      if (selectedVersions.length === 1) return;
      setSelectedVersions(selectedVersions.filter((v) => v !== ver));
    } else {
      setSelectedVersions([...selectedVersions, ver]);
    }
  };

// 1. 단일 본문 생성 핸들러 (로컬은 자동 열기, 웹은 즉시 브라우저 다운로드)
  const handleGeneratePpt = async () => {
    try {
      const params = new URLSearchParams({
        book: book.trim(),
        chapter: String(chapter),
        font_name: fontName,
      });

      if (startVerse && String(startVerse).trim() !== "") {
        params.append("start_verse", String(startVerse).trim());
      }
      if (endVerse && String(endVerse).trim() !== "") {
        params.append("end_verse", String(endVerse).trim());
      }

      selectedVersions.forEach((v) => params.append("versions", v));

      const response = await fetch(`/api/open-ppt?${params.toString()}`, {
        method: "POST",
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => null);
        throw new Error(errorData?.detail || `서버 에러 (${response.status})`);
      }

      // 서버 응답 분기 (JSON이면 로컬 열기 완료, Blob 스트림이면 웹 브라우저 다운로드)
      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        await response.json().catch(() => null);
        return; // 로컬 환경: 파워포인트 자동 열림 완료
      }

      // 웹 환경: 브라우저 파일 다운로드 트리거
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const unit = book.trim() === "시편" ? "편" : "장";
      a.download = `${book.trim()}_${chapter}${unit}.pptx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      alert("PPT 생성 오류: " + err.message);
    }
  };
      // 서버 응답이 JSON(로컬 실행 시)인지 파일 스트림(웹 배포 시)인지 분기
      const contentType = res.headers.get("content-type");
      if (contentType && contentType.includes("application/json")) {
        const data = await res.json();
        // 로컬 환경: 파워포인트 자동 열림 완료
        return;
      }

      // 웹 환경: 브라우저 다운로드 실행
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${selectedBook}_${chapter}장.pptx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      alert(`다중 PPT 생성 오류: ${err.message}`);
    }
  };

  // 2. 여러 본문 한 번에 연속 생성 핸들러
  const handleGenerateMultiPpt = async () => {
    if (!multiQuery.trim()) {
      alert("구절을 입력해 주세요. (예: 창 1:1-3, 요 3:16, 롬 8:28)");
      return;
    }
    try {
      const res = await fetch("/api/open-multi-ppt", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          query_text: multiQuery,
          versions: selectedVersions,
          font_name: fontName,
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "다중 본문 생성에 실패했습니다.");
      }
      await res.json().catch(() => null);
    } catch (err) {
      alert("다중 PPT 생성 오류: " + err.message);
    }
  };

  // 3. 키워드 검색 핸들러
  const handleSearch = async () => {
    if (!searchWord || searchWord.trim().length < 2) {
      alert("검색어를 2글자 이상 입력해 주세요.");
      return;
    }

    setIsSearching(true);
    try {
      const res = await fetch(`/api/search-keyword?keyword=${encodeURIComponent(searchWord.trim())}&version=${encodeURIComponent(selectedVersions[0] || "개역개정")}`);
      if (!res.ok) {
        const errJson = await res.json().catch(() => null);
        throw new Error(errJson?.detail || `검색 오류 (${res.status})`);
      }
      const data = await res.json();
      if (Array.isArray(data)) {
        setSearchResults(data);
        if (data.length === 0) {
          alert("검색 결과가 없습니다.");
        }
      } else {
        setSearchResults([]);
      }
    } catch (err) {
      alert("검색 실패: " + err.message);
      setSearchResults([]);
    } finally {
      setIsSearching(false);
    }
  };

  // 4. 구절 내용 수정 핸들러
  const handleSaveVerse = async () => {
    if (!editVerse) return;
    try {
      const res = await fetch("/api/verse", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          version: selectedVersions[0] || "개역개정",
          book: editVerse.book,
          chapter: editVerse.chapter,
          verse_num: editVerse.verse_num,
          text: editText,
        }),
      });
      if (!res.ok) throw new Error("저장에 실패했습니다.");
      alert("성경 DB가 성공적으로 수정되었습니다!");
      setEditVerse(null);
      handleSearch();
    } catch (err) {
      alert("수정 저장 오류: " + err.message);
    }
  };
// 검색 결과에서 클릭한 특정 구절 즉시 PPT 열기
  const handleQuickPpt = async (item) => {
    try {
      const params = new URLSearchParams({
        book: item.book,
        chapter: String(item.chapter),
        start_verse: String(item.verse_num),
        end_verse: String(item.verse_num),
        font_name: fontName,
      });
      selectedVersions.forEach((v) => params.append("versions", v));

      const response = await fetch(`/api/open-ppt?${params.toString()}`, { method: "POST" });
      if (!response.ok) {
        const errorData = await response.json().catch(() => null);
        throw new Error(errorData?.detail || `서버 에러 (${response.status})`);
      }
      await response.json().catch(() => null);
    } catch (err) {
      alert("PPT 생성 오류: " + err.message);
    }
  };
const handleShutdown = async () => {
    try {
      await fetch("/api/shutdown", { method: "POST" });
    } catch (e) {
      // 프로세스 종료로 인한 통신 단절 무시
    }
  };
  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 p-6 flex flex-col">
      {/* 상단 헤더 */}
      <header className="flex justify-between items-center mb-6 border-b border-slate-700 pb-4">
        <div>
          <h1 className="text-2xl font-extrabold text-amber-400 flex items-center gap-2">
            📖 BibleScript Pro
          </h1>
          <p className="text-xs text-slate-400">교회 방송실 및 예배용 고화질 성경 PPT 자동 생성기</p>
        </div>
        <button
          onClick={handleShutdown}
          className="px-3 py-1.5 bg-red-900/60 hover:bg-red-800 text-red-200 border border-red-700 rounded text-xs transition"
        >
          프로그램 종료
        </button>
      </header>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6 flex-1">
        {/* 좌측: 구절 및 번역본 설정 패널 */}
        <div className="bg-slate-800 p-5 rounded-xl border border-slate-700 space-y-4">
          <h2 className="text-base font-bold text-amber-300">구절 및 번역본 설정</h2>

          <div>
            <label className="block text-xs font-bold text-slate-400 mb-1">번역본 선택 (2개 선택 시 대조 슬라이드)</label>
            <div className="flex flex-wrap gap-1.5">
              {["개역개정", "표준새번역", "NIV"].map((ver) => (
                <button
                  key={ver}
                  type="button"
                  onClick={() => toggleVersion(ver)}
                  className={`px-3 py-1 rounded text-xs font-bold transition ${
                    selectedVersions.includes(ver)
                      ? "bg-amber-500 text-slate-950 shadow"
                      : "bg-slate-700 text-slate-300 hover:bg-slate-600"
                  }`}
                >
                  {ver}
                </button>
              ))}
            </div>
          </div>

          {/* 단일 본문 입력 구역 */}
          <div className="pt-2 border-t border-slate-700">
            <span className="text-xs font-semibold text-slate-400 mb-2 block">단일 본문 입력</span>
            <div className="grid grid-cols-2 gap-2 mb-2">
              <div>
                <label className="block text-[11px] text-slate-400 mb-0.5">성경(약어 가능)</label>
                <input
                  type="text"
                  value={book}
                  placeholder="예: 창, 계"
                  onChange={(e) => setBook(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded p-2 text-sm text-white"
                />
              </div>
              <div>
                <label className="block text-[11px] text-slate-400 mb-0.5">장</label>
                <input
                  type="number"
                  value={chapter}
                  onChange={(e) => setChapter(Number(e.target.value))}
                  className="w-full bg-slate-900 border border-slate-700 rounded p-2 text-sm text-white"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2 mb-3">
              <div>
                <label className="block text-[11px] text-slate-400 mb-0.5">시작 절</label>
                <input
                  type="text"
                  value={startVerse}
                  placeholder="전체"
                  onChange={(e) => setStartVerse(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded p-2 text-sm text-white"
                />
              </div>
              <div>
                <label className="block text-[11px] text-slate-400 mb-0.5">끝 절</label>
                <input
                  type="text"
                  value={endVerse}
                  placeholder="전체"
                  onChange={(e) => setEndVerse(e.target.value)}
                  className="w-full bg-slate-900 border border-slate-700 rounded p-2 text-sm text-white"
                />
              </div>
            </div>

            <button
              type="button"
              onClick={handleGeneratePpt}
              className="w-full py-2.5 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg shadow transition text-sm"
            >
              단일 본문 PPT 생성 및 열기
            </button>
          </div>

          {/* ▼▼▼ 여러 본문 한 번에 입력하는 구역 ▼▼▼ */}
          <div className="pt-3 border-t border-slate-700">
            <label className="block text-xs font-bold text-emerald-400 mb-1">
              ✨ 여러 본문 한 번에 연속 생성
            </label>
            <textarea
              rows={3}
              placeholder="예: 창 1:1-3, 요 3:16, 롬 8:28&#13;&#10;(쉼표, 세미콜론, 줄바꿈으로 구분)"
              value={multiQuery}
              onChange={(e) => setMultiQuery(e.target.value)}
              className="w-full bg-slate-900 border border-slate-700 rounded p-2 text-xs text-white mb-2 focus:border-emerald-500 focus:outline-none"
            />
            <button
              type="button"
              onClick={handleGenerateMultiPpt}
              className="w-full py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-lg shadow transition text-sm"
            >
              여러 본문 합본 PPT 생성 및 즉시 열기
            </button>
          </div>
        </div>

        {/* 우측: 성경 본문 검색 및 오탈자 관리 */}
        <div className="bg-slate-800 p-5 rounded-xl border border-slate-700 md:col-span-2 flex flex-col">
          <h2 className="text-base font-bold text-amber-300 mb-2">성경 본문 검색 및 오탈자 관리</h2>
          <div className="flex gap-2 mb-3">
            <input
              type="text"
              placeholder="검색할 단어 입력 (예: 태초에, 사랑)"
              value={searchWord}
              onChange={(e) => setSearchWord(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") handleSearch(); }}
              className="flex-1 bg-slate-900 border border-slate-700 rounded p-2 text-sm text-white"
            />
            <button
              type="button"
              onClick={handleSearch}
              disabled={isSearching}
              className="px-4 py-2 bg-slate-700 hover:bg-slate-600 rounded font-bold text-sm text-white transition"
            >
              {isSearching ? "검색 중..." : "검색"}
            </button>
          </div>

          <div className="flex-1 max-h-96 overflow-y-auto space-y-2 pr-1">
            {Array.isArray(searchResults) && searchResults.map((r, idx) => (
              <div
                key={idx}
                className="p-3 bg-slate-900/80 border border-slate-800 rounded flex justify-between items-start text-xs"
              >
                <div>
                  <span className="font-bold text-amber-400 mr-2">
                    {r.book} {r.chapter}장 {r.verse_num}절
                  </span>
                  <p className="text-slate-300 mt-1 leading-relaxed">{r.text}</p>
                </div>
                <button
                  type="button"
                  onClick={() => handleQuickPpt(r)}
                 className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-[11px] font-bold rounded text-white shadow transition"
                 title="이 구절로 즉시 PPT 열기"
                 >
                 PPT 열기
                </button>
               <button
                 type="button"
                 onClick={() => {
                      setEditVerse(r);
                      setEditText(r.text);
                 }}
                 className="px-2 py-1 bg-slate-800 hover:bg-slate-700 border border-slate-600 text-[11px] rounded text-slate-300"
             >
                수정
               </button>
              </div>
            ))}
          </div>

          {/* 구절 수정 창 */}
          {editVerse && (
            <div className="mt-3 p-3 bg-slate-950 border border-amber-500/50 rounded-lg">
              <h3 className="text-xs font-bold text-amber-400 mb-1">
                구절 오탈자 수정 및 수동 분할 ({editVerse.book} {editVerse.chapter}장 {editVerse.verse_num}절)
              </h3>
              <textarea
                value={editText}
                onChange={(e) => setEditText(e.target.value)}
                rows={2}
                className="w-full bg-slate-900 border border-slate-700 rounded p-2 text-xs text-slate-200 focus:outline-none"
              />
              <p className="text-[11px] text-slate-400 mt-1">
                * 문장 중간에 <span className="text-amber-400 font-bold">//</span> 기호를 넣으면 해당 위치에서 슬라이드가 나뉩니다.
              </p>
              <div className="flex gap-2 justify-end mt-2">
                <button
                  type="button"
                  onClick={() => setEditVerse(null)}
                  className="px-3 py-1 bg-slate-800 hover:bg-slate-700 text-xs rounded text-slate-300"
                >
                  취소
                </button>
                <button
                  type="button"
                  onClick={handleSaveVerse}
                  className="px-3 py-1 bg-amber-600 hover:bg-amber-500 font-bold text-xs text-white rounded"
                >
                  DB 저장
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}