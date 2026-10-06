@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   rainfield-im 변경사항을 클라우드에 올립니다
echo ============================================
echo.
git add .
git commit -m "수정 %date% %time%"
if errorlevel 1 (
  echo.
  echo [알림] 새로 바뀐 내용이 없어요. (이미 최신)
) else (
  git push
  echo.
  echo [완료] 업로드 끝! 잠시 후 클라우드 앱에 반영됩니다.
  echo        (안 바뀌면 Streamlit 앱에서 우측 ⋮ → Reboot app)
)
echo.
echo 이 창은 아무 키나 누르면 닫힙니다.
pause >nul
