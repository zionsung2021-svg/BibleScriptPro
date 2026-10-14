import React, { useState, useEffect } from "react";

export default function App() {
  const [versions, setVersions] = useState(["개역개정"]);
  const [selectedVersions, setSelectedVersions] = useState(["개역개정"]);
  const [book, setBook] = useState("창세기");
  const [chapter, setChapter] = useState(1);
  const [startVerse, setStartVerse] = useState("");
  const [endVerse, setEndVerse] = useState("");
  const [fontName, setFontName] = useState("경기천년제목 Bold");

  // 자동 업데이트 상태
  const [updateInfo, setUpdateInfo] = useState(null);
  const [isUpdating, setIsUpdating] = useState(false);

  // 여러 본문 일괄 입력 상태
  const [multiQuery, setMultiQuery] = useState("");

  // 키워드 검색 상태
  const [searchWord, setSearchWord] = useState("");
  const [searchResults, setSearchResults] = useState([]);
  const [isSearching, setIsSearching] = useState(false);

  // 구절 인라인 수정 상태
  const [editVerse, setEditVerse] = useState(null);
  const [editText, setEditText] = useState("");

  // 모니터 목록 상태
  const [monitors, setMonitors] = useState([]);
  const [selectedMonitor, setSelectedMonitor] = useState(1);

  // 1. 번역본 목록 불러오기
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

 // 모니터 목록 불러오기
  useEffect(() => {
    fetch("/api/monitors")
      .then((res) => res.json())
      .then((data) => {
        if (data && data.monitors && data.monitors.length > 0) {
          setMonitors(data.monitors);
          // 빔프로젝터나 확장 모니터(!is_primary)가 있으면 그것을 기본 송출 모니터로 자동 지정
          const secondary = data.monitors.find((m) => !m.is_primary);
          if (secondary) {
            setSelectedMonitor(secondary.id);
          } else {
            setSelectedMonitor(data.monitors[0].id);
          }
        }
      })
      .catch((err) => console.error("모니터 감지 실패:", err));
  }, []);

  // 3. 업데이트 확인 로직
  useEffect(() => {
    fetch("/api/update/check")
      .then((res) => res.json())
      .then((data) => {
        if (data && data.has_update) {
          setUpdateInfo(data);
        }
      })
      .catch((err) => console.error("업데이트 확인 실패:", err));
  }, []);

  // 업데이트 실행 핸들러
  const handleApplyUpdate = async () => {
    if (!updateInfo) return;
    if (
      window.confirm(
        `새 버전(v${updateInfo.latest_version})으로 즉시 업데이트하고 프로그램을 재시작할까요?`
      )
    ) {
      setIsUpdating(true);
      try {
        const res = await fetch("/api/update/apply", { method: "POST" });
        if (!res.ok) throw new Error("업데이트 요청 실패");
      } catch (err) {
        alert("업데이트 적용 중 오류가 발생했습니다: " + err.message);
        setIsUpdating(false);
      }
    }
  };

  // 번역본 선택 토글
  const toggleVersion = (ver) => {
    if (selectedVersions.includes(ver)) {
      if (selectedVersions.length === 1) return;
      setSelectedVersions(selectedVersions.filter((v) => v !== ver));
    } else {
      setSelectedVersions([...selectedVersions, ver]);
    }
  };

  // 1. 단일 본문 생성 핸들러
  const handleGeneratePpt = async () => {
    try {
      const params = new URLSearchParams({
        book: book.trim(),
        chapter: String(chapter),
        font_name: fontName,
        target_monitor: String(selectedMonitor),
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

      const contentType = response.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        await response.json().catch(() => null);
        return;
      }

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

  // 2. 여러 본문 연속 생성 핸들러
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
          lines: multiQuery,
          versions: selectedVersions,
          font_name: fontName,
          target_monitor: selectedMonitor,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => null);
        throw new Error(err?.detail || "다중 본문 생성에 실패했습니다.");
      }

      const contentType = res.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        await res.json().catch(() => null);
        return;
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `성경본문_합본.pptx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      alert("다중 PPT 생성 오류: " + err.message);
    }
  };

  // 키워드 검색 핸들러
  const handleSearch = async () => {
    if (!searchWord.trim() || searchWord.trim().length < 2) {
      alert("검색어를 2글자 이상 입력하세요.");
      return;
    }
    setIsSearching(true);
    try {
      const res = await fetch(
        `/api/search-keyword?keyword=${encodeURIComponent(
          searchWord.trim()
        )}&version=${encodeURIComponent(selectedVersions[0] || "개역개정")}`
      );
      if (res.ok) {
        const data = await res.json();
        setSearchResults(data);
      }
    } catch (e) {
      console.error(e);
    } finally {
      setIsSearching(false);
    }
  };

  // 인라인 본문 저장 핸들러
  const handleSaveVerse = async () => {
    if (!editVerse) return;
    try {
      const res = await fetch("/api/verse", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          version: editVerse.version,
          book: editVerse.book,
          chapter: editVerse.chapter,
          verse_num: editVerse.verse_num,
          text: editText,
        }),
      });
      if (res.ok) {
        alert("수정되었습니다.");
        setSearchResults(
          searchResults.map((r) =>
            r.version === editVerse.version &&
            r.book === editVerse.book &&
            r.chapter === editVerse.chapter &&
            r.verse_num === editVerse.verse_num
              ? { ...r, text: editText }
              : r
          )
        );
        setEditVerse(null);
      } else {
        alert("저장에 실패했습니다.");
      }
    } catch (e) {
      alert("오류가 발생했습니다.");
    }
  };

  const handleShutdown = async () => {
    if (window.confirm("프로그램을 완전히 종료하시겠습니까?")) {
      try {
        await fetch("/api/shutdown", { method: "POST" });
      } catch (err) {
        console.error("Shutdown call error:", err);
      }
      window.close();
      document.body.innerHTML = `
        <div style="display:flex;justify-content:center;align-items:center;height:100vh;background:#0d1117;color:#fff;font-family:sans-serif;flex-direction:column;">
          <h2 style="font-size:1.8rem;margin-bottom:12px;">프로그램이 안전하게 종료되었습니다.</h2>
          <p style="color:#8b949e;">이 창을 닫아주세요.</p>
        </div>
      `;
    }
  };

  return (
    <div className="min-h-screen bg-slate-900 text-gray-100 p-6 flex flex-col items-center">
      {/* 상단 헤더: 타이틀 + 업데이트 버튼 + 프로그램 종료 버튼 */}
      <header className="w-full max-w-6xl flex justify-between items-center mb-6">
        <div>
          <h1 className="text-2xl font-bold text-white flex items-center gap-2">
            📖 BibleScript Pro
          </h1>
          <p className="text-sm text-gray-400">예배용 고화질 성경 PPT 자동 생성기</p>
        </div>

        <div className="flex items-center gap-2">
          {/* 업데이트 알림 버튼 (새 버전 감지 시 노출) */}
          {updateInfo && (
            <button
              onClick={handleApplyUpdate}
              disabled={isUpdating}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold text-amber-300 bg-amber-950/80 hover:bg-amber-900 border border-amber-500/70 rounded-lg transition-all animate-pulse shadow-md cursor-pointer disabled:opacity-50"
              title="새 버전을 다운받아 자동으로 적용하고 재시작합니다."
            >
              🚀 {isUpdating ? "업데이트 적용 중..." : `v${updateInfo.latest_version} 업데이트`}
            </button>
          )}

          {/* 프로그램 종료 버튼 */}
          <button
            onClick={handleShutdown}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-rose-300 bg-rose-950/50 hover:bg-rose-900/80 border border-rose-500/40 rounded-lg transition-all duration-200 shadow-sm cursor-pointer"
            title="프로그램 및 백엔드 서버 종료"
          >
            <svg
              xmlns="http://www.w3.org/2000/svg"
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M18.36 6.64a9 9 0 1 1-12.73 0" />
              <line x1="12" y1="2" x2="12" y2="12" />
            </svg>
            프로그램 종료
          </button>
        </div>
      </header>

      <div className="w-full max-w-6xl grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* 좌측: 구절 및 번역본 설정 */}
        <div className="bg-slate-800/80 border border-slate-700/80 rounded-2xl p-6 shadow-xl flex flex-col gap-5">
          <h2 className="text-lg font-bold text-amber-400 border-b border-slate-700 pb-3">
            구절 및 번역본 설정
          </h2>

          {/* 번역본 선택 버튼 */}
          <div>
            <label className="block text-xs font-semibold text-gray-300 mb-2">
              번역본 선택 (2개 선택 시 대조 슬라이드)
            </label>
            <div className="flex gap-2">
              {versions.map((ver) => {
                const active = selectedVersions.includes(ver);
                return (
                  <button
                    key={ver}
                    onClick={() => toggleVersion(ver)}
                    className={`px-3 py-1.5 rounded-lg text-sm font-semibold transition ${
                      active
                        ? "bg-amber-500 text-slate-900 shadow-md font-bold"
                        : "bg-slate-700 text-gray-300 hover:bg-slate-600"
                    }`}
                  >
                    {ver}
                  </button>
                );
              })}
            </div>
          </div>

          {/* 단일 본문 입력 */}
          <div className="bg-slate-900/60 p-4 rounded-xl border border-slate-700/50 flex flex-col gap-3">
            <span className="text-xs font-bold text-blue-400">단일 본문 입력</span>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs text-gray-400 mb-1">성경 (약어 가능)</label>
                <input
                  type="text"
                  value={book}
                  onChange={(e) => setBook(e.target.value)}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-blue-500"
                />
              </div>
              <div>
                <label className="block text-xs text-gray-400 mb-1">장</label>
                <input
                  type="number"
                  min="1"
                  value={chapter}
                  onChange={(e) => setChapter(Number(e.target.value))}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-blue-500"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs text-gray-400 mb-1">시작 절</label>
                <input
                  type="text"
                  placeholder="전체"
                  value={startVerse}
                  onChange={(e) => setStartVerse(e.target.value)}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-blue-500"
                />
              </div>
              <div>
                <label className="block text-xs text-gray-400 mb-1">끝 절</label>
                <input
                  type="text"
                  placeholder="전체"
                  value={endVerse}
                  onChange={(e) => setEndVerse(e.target.value)}
                  className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-blue-500"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs text-gray-400 mb-1">글꼴 선택</label>
              <select
                value={fontName}
                onChange={(e) => setFontName(e.target.value)}
                className="w-full bg-slate-800 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-blue-500"
              >
                <option value="경기천년제목 Bold">경기천년제목 Bold</option>
                <option value="경기천년제목 Medium">경기천년제목 Medium</option>
                <option value="나눔고딕 Bold">나눔고딕 Bold</option>
                <option value="맑은 고딕">맑은 고딕</option>
              </select>
            </div>

            {/* 듀얼 모니터 감지 드롭다운 */}
            {monitors.length > 1 && (
              <div className="mt-1 pt-2 border-t border-slate-800">
                <label className="block text-xs font-semibold text-emerald-400 mb-1">
                  🖥️ 송출 모니터 선택 (듀얼 모니터 감지됨)
                </label>
                <select
                  value={selectedMonitor}
                  onChange={(e) => setSelectedMonitor(Number(e.target.value))}
                  className="w-full bg-slate-800 border border-emerald-600/50 rounded-lg px-3 py-2 text-sm text-amber-300 focus:outline-none"
                >
                  {monitors.map((m) => (
                    <option key={m.id} value={m.id}>
                      {m.name} ({m.width}x{m.height})
                    </option>
                  ))}
                </select>
              </div>
            )}

            <button
              onClick={handleGeneratePpt}
              className="w-full mt-2 bg-blue-600 hover:bg-blue-500 text-white font-bold py-3 rounded-xl transition shadow-lg shadow-blue-900/40"
            >
              단일 본문 PPT 생성 및 열기
            </button>
          </div>

          {/* 여러 본문 연속 생성 */}
          <div className="bg-slate-900/60 p-4 rounded-xl border border-slate-700/50 flex flex-col gap-2">
            <span className="text-xs font-bold text-emerald-400">✨ 여러 본문 한 번에 연속 생성</span>
            <textarea
              rows={4}
              placeholder={"창 1:1\n요 1:1\n렘 1:1"}
              value={multiQuery}
              onChange={(e) => setMultiQuery(e.target.value)}
              className="w-full bg-slate-800 border border-slate-700 rounded-lg p-3 text-sm text-white font-mono focus:outline-none focus:border-emerald-500"
            />
            <button
              onClick={handleGenerateMultiPpt}
              className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-bold py-3 rounded-xl transition shadow-lg shadow-emerald-900/40"
            >
              여러 본문 합본 PPT 생성 및 즉시 열기
            </button>
          </div>
        </div>

        {/* 우측: 성경 본문 검색 및 오탈자 관리 */}
        <div className="bg-slate-800/80 border border-slate-700/80 rounded-2xl p-6 shadow-xl flex flex-col">
          <h2 className="text-lg font-bold text-amber-400 border-b border-slate-700 pb-3 mb-4">
            성경 본문 검색 및 오탈자 관리
          </h2>

          <div className="flex gap-2 mb-4">
            <input
              type="text"
              placeholder="검색할 단어 입력 (예: 태초에, 사랑)"
              value={searchWord}
              onChange={(e) => setSearchWord(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && handleSearch()}
              className="flex-1 bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm text-white focus:outline-none focus:border-amber-500"
            />
            <button
              onClick={handleSearch}
              className="bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-lg text-sm font-semibold text-white transition"
            >
              {isSearching ? "검색중..." : "검색"}
            </button>
          </div>

          {/* 검색 결과 리스트 */}
          <div className="flex-1 overflow-y-auto max-h-[500px] flex flex-col gap-2 pr-1">
            {searchResults.length === 0 ? (
              <div className="text-center text-gray-500 py-16 text-sm">
                검색된 구절이 없습니다.
              </div>
            ) : (
              searchResults.map((item, idx) => (
                <div
                  key={idx}
                  className="bg-slate-900/70 p-3 rounded-lg border border-slate-700/50 hover:border-slate-600 transition text-sm"
                >
                  <div className="flex justify-between items-center text-xs text-amber-400/90 font-bold mb-1">
                    <span>
                      [{item.version}] {item.book} {item.chapter}장 {item.verse_num}절
                    </span>
                    <button
                      onClick={() => {
                        setEditVerse(item);
                        setEditText(item.text);
                      }}
                      className="text-xs bg-slate-800 hover:bg-slate-700 text-blue-400 px-2 py-0.5 rounded border border-slate-700"
                    >
                      수정
                    </button>
                  </div>
                  <p className="text-gray-300 leading-relaxed">{item.text}</p>
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* 구절 수정 모달 */}
      {editVerse && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-slate-800 border border-slate-700 p-6 rounded-2xl w-full max-w-lg shadow-2xl flex flex-col gap-4">
            <h3 className="font-bold text-amber-400">
              구절 본문 수정: {editVerse.book} {editVerse.chapter}:{editVerse.verse_num} ({editVerse.version})
            </h3>
            <textarea
              rows={4}
              value={editText}
              onChange={(e) => setEditText(e.target.value)}
              className="w-full bg-slate-900 border border-slate-700 rounded-lg p-3 text-sm text-white focus:outline-none focus:border-amber-500"
            />
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setEditVerse(null)}
                className="px-4 py-2 bg-slate-700 hover:bg-slate-600 text-sm font-semibold rounded-lg"
              >
                취소
              </button>
              <button
                onClick={handleSaveVerse}
                className="px-4 py-2 bg-amber-500 hover:bg-amber-400 text-slate-900 text-sm font-bold rounded-lg"
              >
                저장
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}