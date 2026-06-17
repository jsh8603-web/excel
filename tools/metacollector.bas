Attribute VB_Name = "MetaCollector"
Option Explicit

'================================================================
' MetaCollector v1.2.1
'   대상: ActiveWorkbook (읽기 전용 스캔, 파일 수정 금지)
'   산출: 스캔 결과를 한 장짜리 텍스트 리포트로 요약
'================================================================
'
' [사용법]
'
' ── 방법 1. 실사용 (회사 PC, VBA 에디터에서 대화형 실행) ──
'   1) 조사 대상 .xlsx/.xlsm 파일을 Excel에서 열어 ActiveWorkbook으로 만든다.
'   2) Alt+F11 로 VBA 에디터 진입 → 메뉴 [파일]→[파일 가져오기] 로
'      MetaCollector.bas 를 Import (또는 기존 MetaCollector 모듈을 이 파일 내용으로 교체).
'   3) F5 또는 [실행] 메뉴에서 RunCollector 프로시저를 실행.
'   4) 새 워크북이 자동 생성되고 "Report" 시트에 텍스트 리포트가 주입된다.
'      원본 파일은 건드리지 않는다 (읽기 전용 스캔).
'
' ── 방법 2. 자동화 (Python 외부 호출, CI/회귀 테스트) ──
'   run-collector.py / run-collector-all.py / run-collector-messy.py 참조.
'   핵심 흐름:
'     import win32com.client as win32
'     xl = win32.DispatchEx("Excel.Application")
'     xl.Visible = False
'     wb = xl.Workbooks.Open(target_xlsm)
'     # 본 .bas 파일을 VBComponents.Import 로 임시 주입
'     xl.VBE.ActiveVBProject.VBComponents.Import(r"...\MetaCollector.bas")
'     report = xl.Run("MetaCollector.CollectReport")   # 문자열 반환
'     wb.Close(SaveChanges=False)                       # 원본 저장 금지
'   주의: Excel 신뢰 센터에서 "VBA 프로젝트 개체 모델 액세스 허용" 체크 필요.
'
' ── 진입점 요약 ──
'   Sub  RunCollector()    — 대화형. 새 워크북에 Report 시트 생성.
'   Func CollectReport() As String — 외부 호출용. 리포트 문자열만 반환.
'
' ── 리포트 섹션 (출력 순서) ──
'   Header(Protection/CalcMode) → Sheets → ColumnTypes → Names → Tables →
'   AutoFilters → Formulas(블록/함수/R1C1/consistency/3D/외부/volatile) →
'   Shapes → PivotTables → Slicers → PowerQuery → Scenarios → DataModel →
'   Charts → Sparklines → ConditionalFormat → Validation → MergedCells →
'   Comments/Notes → Hyperlinks → Outline → VBAProject → CustomDocProps
'
' ── 출력 예시 ──
'   report-full.txt, report-messy-r3.txt ~ r12.txt 가 동일 폴더에 보관됨.
'================================================================

Private Const VERSION As String = "1.2.1"

' LOG_PATH: 회사/집 무관 동적 경로(%TEMP%). Const(컴파일 상수)는 Environ 불가 → Function.
' 원본(archive/.../MetaCollector.bas)은 D:\projects\... 하드코딩 — 배포본만 패치.
Private Function LogPath() As String
    LogPath = Environ$("TEMP") & "\metacollector_section.log"
End Function

Private Sub LogSection(name As String, phase As String)
    Dim fn As Integer
    fn = FreeFile
    On Error Resume Next
    Open LogPath() For Append As #fn
    Print #fn, Format(Now, "hh:nn:ss") & " " & phase & " " & name
    Close #fn
    On Error GoTo 0
End Sub

'---------- Entry points ----------
' Sub: 대화형 실행 — 새 워크북에 리포트 주입 (회사 PC 실제 사용)
Public Sub RunCollector()
    Dim tgt As Workbook
    Set tgt = ActiveWorkbook
    If tgt Is Nothing Then
        MsgBox "대상 워크북이 없습니다.", vbCritical
        Exit Sub
    End If

    Dim report As String
    report = BuildReport(tgt)
    WriteToNewWorkbook report, tgt.Name
End Sub

' Function: Python/외부 호출용 — 리포트 문자열 반환 (새 워크북 미생성)
Public Function CollectReport() As String
    On Error GoTo errH
    CollectReport = BuildReport(ActiveWorkbook)
    Exit Function
errH:
    CollectReport = "ERROR: " & Err.Number & " " & Err.Description & " @ " & Err.Source
End Function

'---------- Report builder ----------
Private Function BuildReport(wb As Workbook) As String
    ' 로그 초기화
    Dim fn As Integer: fn = FreeFile
    On Error Resume Next
    Open LogPath() For Output As #fn: Close #fn
    On Error GoTo 0

    Dim s As String
    s = Header(wb)
    LogSection "Sheets", "START": s = s & SectionSheets(wb): LogSection "Sheets", "END"
    LogSection "ColumnTypes", "START": s = s & SectionColumnTypes(wb): LogSection "ColumnTypes", "END"
    LogSection "Names", "START": s = s & SectionNames(wb): LogSection "Names", "END"
    LogSection "Tables", "START": s = s & SectionTables(wb): LogSection "Tables", "END"
    LogSection "AutoFilters", "START": s = s & SectionAutoFilters(wb): LogSection "AutoFilters", "END"
    LogSection "Formulas", "START": s = s & SectionFormulas(wb): LogSection "Formulas", "END"
    LogSection "Shapes", "START": s = s & SectionShapes(wb): LogSection "Shapes", "END"
    LogSection "Pivots", "START": s = s & SectionPivots(wb): LogSection "Pivots", "END"
    LogSection "Slicers", "START": s = s & SectionSlicers(wb): LogSection "Slicers", "END"
    LogSection "Queries", "START": s = s & SectionQueries(wb): LogSection "Queries", "END"
    LogSection "Scenarios", "START": s = s & SectionScenarios(wb): LogSection "Scenarios", "END"
    LogSection "DataModel", "START": s = s & SectionDataModel(wb): LogSection "DataModel", "END"
    LogSection "Charts", "START": s = s & SectionCharts(wb): LogSection "Charts", "END"
    LogSection "Sparklines", "START": s = s & SectionSparklines(wb): LogSection "Sparklines", "END"
    LogSection "CF", "START": s = s & SectionConditionalFormat(wb): LogSection "CF", "END"
    LogSection "Validation", "START": s = s & SectionValidation(wb): LogSection "Validation", "END"
    LogSection "MergedCells", "START": s = s & SectionMergedCells(wb): LogSection "MergedCells", "END"
    LogSection "Comments", "START": s = s & SectionComments(wb): LogSection "Comments", "END"
    LogSection "Hyperlinks", "START": s = s & SectionHyperlinks(wb): LogSection "Hyperlinks", "END"
    LogSection "Grouping", "START": s = s & SectionGrouping(wb): LogSection "Grouping", "END"
    LogSection "VBA", "START": s = s & SectionVBAProject(wb): LogSection "VBA", "END"
    LogSection "DocProps", "START": s = s & SectionDocProperties(wb): LogSection "DocProps", "END"
    s = s & Footer()
    BuildReport = s
End Function

Private Function Header(wb As Workbook) As String
    Dim s As String
    Dim sizeKB As String: sizeKB = "?"
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(wb.FullName) Then
        sizeKB = Format(fso.GetFile(wb.FullName).Size \ 1024, "#,##0") & " KB"
    End If
    On Error GoTo 0

    Dim protTag As String: protTag = ""
    On Error Resume Next
    If wb.ProtectStructure Then protTag = protTag & "structure "
    If wb.ProtectWindows Then protTag = protTag & "windows "
    On Error GoTo 0
    If Len(protTag) = 0 Then protTag = "none"

    Dim calcMode As String
    On Error Resume Next
    Select Case Application.Calculation
        Case -4105: calcMode = "automatic"
        Case -4135: calcMode = "manual"
        Case 2: calcMode = "semiautomatic"
        Case Else: calcMode = "?"
    End Select
    On Error GoTo 0

    s = "=== MetaCollector v" & VERSION & " ===" & vbCrLf
    s = s & "File       : " & wb.Name & vbCrLf
    s = s & "Path       : " & MaskPath(wb.path) & vbCrLf
    s = s & "Size       : " & sizeKB & vbCrLf
    s = s & "Generated  : " & Format(Now, "yyyy-mm-dd hh:nn") & vbCrLf
    s = s & "Sheets     : " & wb.Sheets.Count & vbCrLf
    s = s & "Protection : " & Trim(protTag) & vbCrLf
    s = s & "CalcMode   : " & calcMode & vbCrLf & vbCrLf
    Header = s
End Function

Private Function Footer() As String
    Footer = vbCrLf & "=== END ===" & vbCrLf
End Function

'---------- Sheets ----------
Private Function SectionSheets(wb As Workbook) As String
    Dim s As String, sh As Object, vis As String, usedR As Range, rows As Long, cols As Long
    s = "[Sheets]" & vbCrLf
    For Each sh In wb.Sheets
        Select Case sh.Visible
            Case xlSheetVisible: vis = "visible"
            Case xlSheetHidden: vis = "HIDDEN"
            Case xlSheetVeryHidden: vis = "VERY_HIDDEN"
            Case Else: vis = "?"
        End Select
        If TypeName(sh) = "Worksheet" Then
            On Error Resume Next
            Set usedR = sh.UsedRange
            rows = 0: cols = 0
            If Not usedR Is Nothing Then
                rows = usedR.rows.Count
                cols = usedR.Columns.Count
            End If
            Dim protTag As String: protTag = ""
            If sh.ProtectContents Then protTag = ", protected"
            ' 숨긴 행/열 카운트 (used range 내)
            Dim hiddenR As Long, hiddenC As Long, hi As Long
            hiddenR = 0: hiddenC = 0
            If Not usedR Is Nothing And rows > 0 Then
                For hi = 1 To rows
                    If usedR.rows(hi).Hidden Then hiddenR = hiddenR + 1
                Next hi
                For hi = 1 To cols
                    If usedR.Columns(hi).Hidden Then hiddenC = hiddenC + 1
                Next hi
            End If
            On Error GoTo 0
            Dim hideTag As String: hideTag = ""
            If hiddenR > 0 Then hideTag = hideTag & " hidden_rows=" & hiddenR
            If hiddenC > 0 Then hideTag = hideTag & " hidden_cols=" & hiddenC
            s = s & "  - " & sh.Name & " [" & vis & protTag & "] used=" & rows & "x" & cols & hideTag & vbCrLf
        Else
            s = s & "  - " & sh.Name & " [" & vis & "] type=" & TypeName(sh) & vbCrLf
        End If
    Next sh
    SectionSheets = s & vbCrLf
End Function

'---------- Named Ranges ----------
Private Function SectionNames(wb As Workbook) As String
    Dim s As String, nm As Name, ref As String, isLambda As Boolean
    s = "[Named Ranges]" & vbCrLf
    If wb.Names.Count = 0 Then
        s = s & "  (none)" & vbCrLf & vbCrLf
        SectionNames = s
        Exit Function
    End If
    Dim nmEmitted As Long: nmEmitted = 0
    For Each nm In wb.Names
        ' 내부 Excel 등록명 (_xlfn.*, _xlpm.*) 스킵
        If Left(nm.Name, 5) = "_xlfn" Or Left(nm.Name, 5) = "_xlpm" Then GoTo nxt
        On Error Resume Next
        ref = nm.RefersTo
        On Error GoTo 0
        Dim hiddenTag As String: hiddenTag = ""
        On Error Resume Next
        If nm.Visible = False Then hiddenTag = " [hidden]"
        On Error GoTo 0
        isLambda = (InStr(ref, "LAMBDA(") > 0)
        If isLambda Then
            s = s & "  - " & nm.Name & hiddenTag & "  [LAMBDA]  " & Truncate(ref, 120) & vbCrLf
        Else
            s = s & "  - " & nm.Name & hiddenTag & "  " & Truncate(MaskExternal(ref), 120) & vbCrLf
        End If
        nmEmitted = nmEmitted + 1
nxt:
    Next nm
    If nmEmitted = 0 Then s = s & "  (none)" & vbCrLf
    SectionNames = s & vbCrLf
End Function

'---------- ListObjects ----------
Private Function SectionTables(wb As Workbook) As String
    Dim s As String, ws As Worksheet, lo As ListObject, cnt As Long
    s = "[Tables (ListObjects)]" & vbCrLf
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            cnt = cnt + 1
            s = s & "  - " & ws.Name & "!" & lo.Name & "  range=" & lo.Range.Address(False, False) _
                & "  rows=" & lo.ListRows.Count & "  cols=" & lo.ListColumns.Count & vbCrLf
        Next lo
    Next ws
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionTables = s & vbCrLf
End Function

'---------- Formulas: L1 stats + L2 R1C1 pattern grouping + external refs ----------
Private Function SectionFormulas(wb As Workbook) As String
    ' 스케일 전략:
    '   1) ws.UsedRange.SpecialCells(xlCellTypeFormulas) 는 Excel 내부 인덱스라 수만줄에도 즉시 반환.
    '   2) 각 Area 는 빈 행으로 끊긴 수식 덩어리를 자연스럽게 분리한다 ("새 테이블 시작" 패턴).
    '   3) Area.Formula / Area.FormulaR1C1 를 2D Variant 배열로 한 번에 벌크 로드 (COM 왕복 2회/Area).
    '   4) 배열을 VBA 메모리에서 순회 — 셀 단위 COM 없음.
    '   5) 주소는 CellAddr(row, col) 헬퍼로 직접 계산.
    Dim s As String
    Dim ws As Worksheet, fArea As Range, sub_area As Range
    Dim patterns As Object, funcs As Object, externals As Object
    Dim sheetLayouts As Object, threeDRefs As Object
    Dim indirectCnt As Long, offsetCnt As Long
    Dim totalFormulas As Long

    Set patterns = CreateObject("Scripting.Dictionary")
    Set funcs = CreateObject("Scripting.Dictionary")
    Set externals = CreateObject("Scripting.Dictionary")
    Set sheetLayouts = CreateObject("Scripting.Dictionary")
    Set threeDRefs = CreateObject("Scripting.Dictionary")

    For Each ws In wb.Worksheets
        Set fArea = Nothing
        On Error Resume Next
        Set fArea = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If fArea Is Nothing Then GoTo nxtSheetF

        Dim areaIdx As Long: areaIdx = 0
        Dim layout As String: layout = ""

        For Each sub_area In fArea.Areas
            areaIdx = areaIdx + 1
            Dim ar As Long, ac As Long, rr As Long, cc As Long
            ar = sub_area.row
            ac = sub_area.Column
            rr = sub_area.rows.Count
            cc = sub_area.Columns.Count

            layout = layout & "      Blk" & areaIdx & ": " & sub_area.Address(False, False) _
                & " (" & rr & "x" & cc & ")" & vbCrLf

            ' 벌크 로드 — 2D 배열. 단일 셀인 경우 Variant 가 스칼라로 오므로 수동 래핑.
            Dim fArr As Variant, rArr As Variant
            If sub_area.Cells.Count = 1 Then
                ReDim fArr(1 To 1, 1 To 1)
                ReDim rArr(1 To 1, 1 To 1)
                fArr(1, 1) = sub_area.Formula
                rArr(1, 1) = sub_area.FormulaR1C1
            Else
                fArr = sub_area.Formula
                rArr = sub_area.FormulaR1C1
            End If

            Dim ri As Long, ci As Long
            Dim f As String, r1c1 As String, key As String, addr As String
            For ri = 1 To rr
                For ci = 1 To cc
                    f = CStr(fArr(ri, ci))
                    If Len(f) = 0 Then GoTo nxtCellF  ' 사각 Area 안의 빈 셀
                    r1c1 = CStr(rArr(ri, ci))
                    totalFormulas = totalFormulas + 1

                    CountFunctions f, funcs

                    If InStr(f, "[") > 0 And InStr(f, ".xls") > 0 Then
                        Dim extKey As String
                        extKey = ExtractExternalRef(f)
                        If Len(extKey) > 0 Then
                            If Not externals.Exists(extKey) Then externals.Add extKey, 0
                            externals(extKey) = externals(extKey) + 1
                        End If
                    End If

                    If InStr(UCase(f), "INDIRECT(") > 0 Then indirectCnt = indirectCnt + 1
                    If InStr(UCase(f), "OFFSET(") > 0 Then offsetCnt = offsetCnt + 1

                    addr = CellAddr(ar + ri - 1, ac + ci - 1)

                    If Has3DRef(f) Then
                        Dim tdKey As String
                        tdKey = ws.Name & "!" & addr
                        If Not threeDRefs.Exists(tdKey) Then
                            threeDRefs.Add tdKey, f
                        End If
                    End If
                    ' key = sheet || blockTag || r1c1  (블록별 분리 — "끊긴 테이블" 가시화)
                    key = ws.Name & "||Blk" & areaIdx & "||" & r1c1
                    Dim cseFlag As String: cseFlag = "0"
                    On Error Resume Next
                    If sub_area.Cells(ri, ci).HasArray Then cseFlag = "1"
                    On Error GoTo 0
                    If Not patterns.Exists(key) Then
                        patterns.Add key, addr & "|" & addr & "|1|" & cseFlag
                    Else
                        Dim parts() As String
                        parts = Split(patterns(key), "|")
                        Dim curCse As String: curCse = "0"
                        If UBound(parts) >= 3 Then curCse = parts(3)
                        If cseFlag = "1" Then curCse = "1"
                        patterns(key) = parts(0) & "|" & addr & "|" & (CLng(parts(2)) + 1) & "|" & curCse
                    End If
nxtCellF:
                Next ci
            Next ri
        Next sub_area

        If Len(layout) > 0 And Not sheetLayouts.Exists(ws.Name) Then
            sheetLayouts.Add ws.Name, layout
        End If
nxtSheetF:
        Set fArea = Nothing
    Next ws

    s = "[Formulas]" & vbCrLf
    s = s & "  Total formula cells: " & totalFormulas & vbCrLf & vbCrLf

    ' 블록 레이아웃 — 빈 행으로 끊긴 수식 테이블의 경계를 그대로 노출
    If sheetLayouts.Count > 0 Then
        s = s & "  -- Formula block layout (contiguous areas per sheet) --" & vbCrLf
        Dim sk As Variant
        For Each sk In sheetLayouts.keys
            s = s & "    " & CStr(sk) & ":" & vbCrLf
            s = s & sheetLayouts(sk)
        Next sk
        s = s & vbCrLf
    End If

    ' L1 함수 사용 통계
    s = s & "  -- Function usage (L1) --" & vbCrLf
    s = s & FormatDictSorted(funcs, 30) & vbCrLf

    ' L2 R1C1 패턴 그룹 (블록별 분리, 상위 N개 by 카운트)
    s = s & "  -- R1C1 pattern groups (L2, top 20) --" & vbCrLf
    s = s & FormatPatternsTop(patterns, 20) & vbCrLf

    ' Formula Consistency anomaly — 같은 블록 내에서 다수 패턴 vs 소수 아웃라이어
    ' 탐지 로직: sheet||block 별로 집계, 블록 총 셀수 >= 5 AND 지배 패턴이 >=80% AND 소수 패턴 count == 1 이면 보고
    s = s & "  -- Formula consistency anomalies (outliers in uniform blocks) --" & vbCrLf
    s = s & FormatConsistencyAnomalies(patterns) & vbCrLf

    ' 3D 참조 (Sheet1:Sheet2!range) — top 20 커트와 별개로 노출
    s = s & "  -- 3D references (multi-sheet) --" & vbCrLf
    If threeDRefs.Count = 0 Then
        s = s & "    (none)" & vbCrLf
    Else
        Dim tk As Variant, tdIdx As Long
        tdIdx = 0
        For Each tk In threeDRefs.keys
            tdIdx = tdIdx + 1
            If tdIdx > 10 Then
                s = s & "    ... (+" & (threeDRefs.Count - 10) & " more)" & vbCrLf
                Exit For
            End If
            s = s & "    - " & CStr(tk) & ": " & Truncate(CStr(threeDRefs(tk)), 120) & vbCrLf
        Next tk
    End If
    s = s & vbCrLf

    ' 외부 파일 참조 — externals (수식 스캔) + LinkSources (워크북 레벨) 를 파일명 기준 병합
    s = s & "  -- External file references --" & vbCrLf
    Dim combined As Object
    Set combined = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Dim links As Variant, i As Long
    links = wb.LinkSources(xlExcelLinks)
    If IsArray(links) Then
        For i = LBound(links) To UBound(links)
            Dim linkFile As String, linkMasked As String
            linkFile = MaskFileOnly(CStr(links(i)))
            linkMasked = MaskPath(CStr(links(i)))
            If Not combined.Exists(linkFile) Then
                combined.Add linkFile, linkMasked & "|0"
            End If
        Next i
    End If
    On Error GoTo 0

    Dim ek As Variant
    For Each ek In externals.keys
        Dim eFile As String
        eFile = MaskFileOnly(CStr(ek))
        If combined.Exists(eFile) Then
            Dim cp() As String
            cp = Split(combined(eFile), "|")
            combined(eFile) = cp(0) & "|" & externals(ek)
        Else
            combined.Add eFile, eFile & "|" & externals(ek)
        End If
    Next ek

    If combined.Count = 0 Then
        s = s & "    (none)" & vbCrLf
    Else
        Dim ck As Variant
        For Each ck In combined.keys
            Dim parts2() As String
            parts2 = Split(combined(ck), "|")
            s = s & "    - " & parts2(0)
            If CLng(parts2(1)) > 0 Then
                s = s & " (in formulas: x" & parts2(1) & ")"
            End If
            s = s & vbCrLf
        Next ck
    End If
    s = s & vbCrLf

    ' Volatile flag
    s = s & "  -- Volatile/dependency-breaking functions --" & vbCrLf
    s = s & "    INDIRECT: " & indirectCnt & vbCrLf
    s = s & "    OFFSET  : " & offsetCnt & vbCrLf & vbCrLf

    SectionFormulas = s
End Function

'---------- PivotTables ----------
Private Function SectionPivots(wb As Workbook) As String
    Dim s As String, ws As Worksheet, pt As PivotTable, pf As PivotField, cnt As Long
    s = "[PivotTables]" & vbCrLf
    For Each ws In wb.Worksheets
        On Error Resume Next
        For Each pt In ws.PivotTables
            cnt = cnt + 1
            s = s & "  - " & ws.Name & "!" & pt.Name & vbCrLf
            s = s & "      source: " & Truncate(MaskPath(CStr(pt.SourceData)), 100) & vbCrLf
            s = s & "      rows  : " & JoinFields(pt.RowFields) & vbCrLf
            s = s & "      cols  : " & JoinFields(pt.ColumnFields) & vbCrLf
            s = s & "      data  : " & JoinFields(pt.DataFields) & vbCrLf
            s = s & "      filter: " & JoinFields(pt.PageFields) & vbCrLf
            ' 계산필드
            Dim cfStr As String: cfStr = ""
            Dim cf As PivotField
            For Each cf In pt.CalculatedFields
                cfStr = cfStr & cf.Name & "=" & cf.Formula & "; "
            Next cf
            If Len(cfStr) > 0 Then s = s & "      calc  : " & cfStr & vbCrLf
        Next pt
        On Error GoTo 0
    Next ws
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionPivots = s & vbCrLf
End Function

'---------- Slicers / Timelines ----------
Private Function SectionSlicers(wb As Workbook) As String
    Dim s As String, sc As SlicerCache, sl As Slicer, cnt As Long, kind As String
    s = "[Slicers & Timelines]" & vbCrLf
    On Error Resume Next
    For Each sc In wb.SlicerCaches
        cnt = cnt + 1
        Select Case sc.SlicerCacheType
            Case 1: kind = "Slicer"
            Case 2: kind = "Timeline"
            Case Else: kind = "type=" & sc.SlicerCacheType
        End Select
        s = s & "  - [" & kind & "] " & sc.Name & "  field=" & sc.SourceName & vbCrLf
    Next sc
    On Error GoTo 0
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionSlicers = s & vbCrLf
End Function

'---------- Power Query ----------
' 쿼리당 M code 상한 M_CAP 줄. 초과분은 "... (+N more lines)" 표시.
Private Function SectionQueries(wb As Workbook) As String
    Const M_CAP As Long = 40
    Dim s As String, q As WorkbookQuery, cnt As Long
    s = "[Power Query]" & vbCrLf
    On Error Resume Next
    For Each q In wb.Queries
        cnt = cnt + 1
        s = s & "  - " & q.Name & vbCrLf
        s = s & "    M code:" & vbCrLf
        Dim mLines() As String, li As Long, total As Long
        mLines = Split(Replace(q.Formula, vbCrLf, vbLf), vbLf)
        total = UBound(mLines) - LBound(mLines) + 1
        For li = LBound(mLines) To UBound(mLines)
            If (li - LBound(mLines)) >= M_CAP Then
                s = s & "      ... (+" & (total - M_CAP) & " more lines)" & vbCrLf
                Exit For
            End If
            s = s & "      " & mLines(li) & vbCrLf
        Next li
        s = s & vbCrLf
    Next q
    On Error GoTo 0
    If cnt = 0 Then s = s & "  (none)" & vbCrLf

    ' Legacy QueryTables (pre-Power-Query, ODBC/SQL/Web/Text imports)
    Dim qtWs As Worksheet, qt As Object, qtCnt As Long
    For Each qtWs In wb.Worksheets
        On Error Resume Next
        For Each qt In qtWs.QueryTables
            qtCnt = qtCnt + 1
            If qtCnt = 1 Then s = s & vbCrLf & "  -- Legacy QueryTables --" & vbCrLf
            s = s & "    - " & qtWs.Name & "!" & qt.Name & vbCrLf
            Dim qConn As String, qCmd As String, qCmdType As String
            qConn = "": qCmd = "": qCmdType = ""
            qConn = CStr(qt.Connection)
            qCmd = CStr(qt.CommandText)
            qCmdType = CStr(qt.CommandType)
            If Len(qConn) > 0 Then _
                s = s & "      connection: " & Truncate(MaskExternal(qConn), 200) & vbCrLf
            If Len(qCmd) > 0 Then _
                s = s & "      command: " & Truncate(qCmd, 200) & "  type=" & qCmdType & vbCrLf
        Next qt
        On Error GoTo 0
    Next qtWs

    ' Connections (with connection-string detail if available)
    Dim conn As WorkbookConnection, ccnt As Long
    On Error Resume Next
    For Each conn In wb.Connections
        ccnt = ccnt + 1
        If ccnt = 1 Then s = s & vbCrLf & "  -- Connections --" & vbCrLf
        s = s & "    - " & conn.Name & "  type=" & conn.Type & vbCrLf
        Dim connStr As String, cmdText As String
        connStr = "": cmdText = ""
        Select Case conn.Type
            Case 1 ' xlConnectionTypeOLEDB
                connStr = CStr(conn.OLEDBConnection.Connection)
                cmdText = CStr(conn.OLEDBConnection.CommandText)
            Case 2 ' xlConnectionTypeODBC
                connStr = CStr(conn.ODBCConnection.Connection)
                cmdText = CStr(conn.ODBCConnection.CommandText)
        End Select
        If Len(connStr) > 0 Then _
            s = s & "      conn: " & Truncate(MaskExternal(connStr), 200) & vbCrLf
        If Len(cmdText) > 0 Then _
            s = s & "      cmd:  " & Truncate(cmdText, 200) & vbCrLf
    Next conn
    On Error GoTo 0
    SectionQueries = s & vbCrLf
End Function

'---------- Charts ----------
Private Function SectionCharts(wb As Workbook) As String
    Dim s As String, ws As Worksheet, co As ChartObject, cnt As Long
    s = "[Charts]" & vbCrLf
    For Each ws In wb.Worksheets
        On Error Resume Next
        For Each co In ws.ChartObjects
            cnt = cnt + 1
            Dim pivTag As String: pivTag = ""
            Dim plTest As Object
            Set plTest = Nothing
            Set plTest = co.Chart.PivotLayout
            If Not plTest Is Nothing Then pivTag = " [PivotChart]"
            Set plTest = Nothing
            s = s & "  - " & ws.Name & "!" & co.Name & "  type=" & ChartTypeName(co.Chart.ChartType) & pivTag & vbCrLf
        Next co
        On Error GoTo 0
    Next ws
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionCharts = s & vbCrLf
End Function

'---------- Conditional Formatting (with rule type breakdown) ----------
Private Function SectionConditionalFormat(wb As Workbook) As String
    Dim s As String, ws As Worksheet, n As Long, total As Long, i As Long
    s = "[Conditional Formatting]" & vbCrLf
    For Each ws In wb.Worksheets
        n = 0
        On Error Resume Next
        n = ws.Cells.FormatConditions.Count
        On Error GoTo 0
        If n > 0 Then
            s = s & "  - " & ws.Name & ": " & n & " rules" & vbCrLf
            Dim showLimit As Long: showLimit = n
            If showLimit > 30 Then showLimit = 30
            On Error Resume Next
            For i = 1 To showLimit
                Dim fc As Object, tn As String
                Set fc = ws.Cells.FormatConditions(i)
                tn = CFTypeName(fc.Type)
                Dim appliesTo As String: appliesTo = ""
                appliesTo = fc.appliesTo.Address(False, False)
                Dim variantTag As String: variantTag = ""
                If tn = "IconSet" Then
                    Dim iconId As Long: iconId = 0
                    iconId = fc.IconSet.ID
                    variantTag = " (" & IconSetName(iconId) & ")"
                End If
                s = s & "      [" & i & "] " & tn & variantTag & " @ " & Truncate(appliesTo, 40) & vbCrLf
                Set fc = Nothing
            Next i
            On Error GoTo 0
            If n > showLimit Then s = s & "      ... (+" & (n - showLimit) & " more)" & vbCrLf
            total = total + n
        End If
    Next ws
    If total = 0 Then s = s & "  (none)" & vbCrLf
    SectionConditionalFormat = s & vbCrLf
End Function

Private Function IconSetName(id As Long) As String
    ' XlIconSet enum (대표값만 매핑)
    Select Case id
        Case 1: IconSetName = "3Arrows"
        Case 2: IconSetName = "3ArrowsGray"
        Case 3: IconSetName = "3Flags"
        Case 4: IconSetName = "3TrafficLights1"
        Case 5: IconSetName = "3TrafficLights2"
        Case 6: IconSetName = "3Signs"
        Case 7: IconSetName = "3Symbols"
        Case 8: IconSetName = "3Symbols2"
        Case 9: IconSetName = "4Arrows"
        Case 10: IconSetName = "4ArrowsGray"
        Case 11: IconSetName = "4RedToBlack"
        Case 12: IconSetName = "4Rating"
        Case 13: IconSetName = "4TrafficLights"
        Case 14: IconSetName = "5Arrows"
        Case 15: IconSetName = "5ArrowsGray"
        Case 16: IconSetName = "5Rating"
        Case 17: IconSetName = "5Quarters"
        Case Else: IconSetName = "iconset" & id
    End Select
End Function

Private Function CFTypeName(t As Long) As String
    Select Case t
        Case 1: CFTypeName = "CellValue"
        Case 2: CFTypeName = "Expression"
        Case 3: CFTypeName = "ColorScale"
        Case 4: CFTypeName = "DataBar"
        Case 5: CFTypeName = "Top10"
        Case 6: CFTypeName = "IconSet"
        Case 8: CFTypeName = "UniqueValues"
        Case 9: CFTypeName = "TextString"
        Case 10: CFTypeName = "Blanks"
        Case 11: CFTypeName = "TimePeriod"
        Case 12: CFTypeName = "AboveAverage"
        Case 13: CFTypeName = "NoBlanks"
        Case 16: CFTypeName = "Errors"
        Case 17: CFTypeName = "NoErrors"
        Case Else: CFTypeName = "type" & t
    End Select
End Function

'---------- Data Validation ----------
Private Function SectionValidation(wb As Workbook) As String
    Dim s As String, ws As Worksheet, va As Range, cnt As Long
    Dim wasProtected As Boolean
    s = "[Data Validation]" & vbCrLf
    For Each ws In wb.Worksheets
        ' 보호된 시트는 SpecialCells 실패 — 일시 해제 시도 후 원복
        wasProtected = False
        On Error Resume Next
        If ws.ProtectContents Then
            ws.Unprotect
            If Err.Number = 0 Then wasProtected = True
            Err.Clear
        End If
        Set va = ws.Cells.SpecialCells(xlCellTypeAllValidation)
        On Error GoTo 0
        If Not va Is Nothing Then
            s = s & "  - " & ws.Name & ": " & va.Cells.Count & " cells (" _
                & va.Address(False, False) & ")" & vbCrLf
            cnt = cnt + 1
        End If
        Set va = Nothing
        If wasProtected Then
            On Error Resume Next
            ws.Protect
            On Error GoTo 0
        End If
    Next ws
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionValidation = s & vbCrLf
End Function

Private Function ChartTypeName(t As Long) As String
    Select Case t
        Case 4: ChartTypeName = "Line"
        Case 5: ChartTypeName = "Pie"
        Case 51: ChartTypeName = "ColumnClustered"
        Case 52: ChartTypeName = "ColumnStacked"
        Case -4169: ChartTypeName = "XYScatter"
        Case 15: ChartTypeName = "XYScatter"
        Case 65: ChartTypeName = "BarClustered"
        Case 76: ChartTypeName = "Area"
        Case Else: ChartTypeName = "type" & t
    End Select
End Function

'---------- VBA Project ----------
' - Document 타입 모듈은 comp.Name 이 CodeName (Sheet4 등). 실제 시트명으로 매핑.
' - 시그니처는 모듈당 최대 20개, 초과분은 "... (+M more)" 로 축약.
Private Function SectionVBAProject(wb As Workbook) As String
    Const SIG_CAP As Long = 20
    Dim s As String
    s = "[VBA Project]" & vbCrLf
    On Error Resume Next
    Dim vbp As Object
    Set vbp = wb.VBProject
    If vbp Is Nothing Then
        s = s & "  (access denied or no project)" & vbCrLf & vbCrLf
        SectionVBAProject = s
        Exit Function
    End If

    ' CodeName → 시트명 매핑 구축
    Dim codeNameMap As Object
    Set codeNameMap = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        On Error Resume Next
        codeNameMap.Add ws.CodeName, ws.Name
        On Error GoTo 0
    Next ws

    Dim comp As Object
    For Each comp In vbp.VBComponents
        Dim lineCount As Long
        lineCount = 0
        On Error Resume Next
        lineCount = comp.CodeModule.CountOfLines
        On Error GoTo 0
        ' 빈 Document(시트/통합문서) 모듈은 스킵 — 노이즈 제거
        If comp.Type = 100 And lineCount = 0 Then GoTo nextComp

        Dim displayName As String
        displayName = comp.Name
        If comp.Type = 100 And codeNameMap.Exists(comp.Name) Then
            displayName = comp.Name & " (" & codeNameMap(comp.Name) & ")"
        End If

        s = s & "  - " & displayName & "  type=" & ComponentTypeName(comp.Type) & "  lines=" & lineCount & vbCrLf

        ' Sub/Function 시그니처 추출 — 상한 SIG_CAP
        Dim i As Long, line As String, sigCount As Long
        sigCount = 0
        For i = 1 To lineCount
            line = comp.CodeModule.Lines(i, 1)
            Dim trimmed As String
            trimmed = Trim(line)
            If InStr(1, trimmed, "Sub ", vbTextCompare) = 1 _
               Or InStr(1, trimmed, "Function ", vbTextCompare) = 1 _
               Or InStr(1, trimmed, "Public Sub ", vbTextCompare) = 1 _
               Or InStr(1, trimmed, "Public Function ", vbTextCompare) = 1 _
               Or InStr(1, trimmed, "Private Sub ", vbTextCompare) = 1 _
               Or InStr(1, trimmed, "Private Function ", vbTextCompare) = 1 Then
                sigCount = sigCount + 1
                If sigCount <= SIG_CAP Then
                    s = s & "      " & Truncate(trimmed, 100) & vbCrLf
                End If
            End If
        Next i
        If sigCount > SIG_CAP Then
            s = s & "      ... (+" & (sigCount - SIG_CAP) & " more)" & vbCrLf
        End If
nextComp:
    Next comp
    On Error GoTo 0
    SectionVBAProject = s & vbCrLf
End Function

Private Function ComponentTypeName(t As Long) As String
    Select Case t
        Case 1: ComponentTypeName = "StdModule"
        Case 2: ComponentTypeName = "ClassModule"
        Case 3: ComponentTypeName = "UserForm"
        Case 100: ComponentTypeName = "Document"
        Case Else: ComponentTypeName = "type" & t
    End Select
End Function

'---------- Document Properties ----------
Private Function SectionDocProperties(wb As Workbook) As String
    Dim s As String, p As Object, cnt As Long
    s = "[Custom Document Properties]" & vbCrLf
    On Error Resume Next
    For Each p In wb.CustomDocumentProperties
        cnt = cnt + 1
        s = s & "  - " & p.Name & " = " & Truncate(CStr(p.Value), 100) & vbCrLf
    Next p
    On Error GoTo 0
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionDocProperties = s & vbCrLf
End Function

'================================================================
' Helpers
'================================================================

Private Sub CountFunctions(f As String, dict As Object)
    ' 매우 단순한 함수명 추출 — 대문자 단어 직후 "(" 패턴
    Dim i As Long, ch As String, buf As String
    Dim upF As String: upF = UCase(f)
    For i = 1 To Len(upF)
        ch = Mid(upF, i, 1)
        If (ch >= "A" And ch <= "Z") Or ch = "_" Or (ch >= "0" And ch <= "9" And Len(buf) > 0) Or ch = "." Then
            buf = buf & ch
        Else
            If ch = "(" And Len(buf) >= 2 Then
                ' _xlfn. 접두사 제거
                If Left(buf, 6) = "_XLFN." Then buf = Mid(buf, 7)
                If Left(buf, 10) = "_XLFN._XLWS." Then buf = Mid(buf, 11)
                If Not dict.Exists(buf) Then dict.Add buf, 0
                dict(buf) = dict(buf) + 1
            End If
            buf = ""
        End If
    Next i
End Sub

Private Function ExtractExternalRef(f As String) As String
    ' ='[파일명.xlsx]시트'!$A$1 같은 패턴에서 [파일명.xlsx] 부분 추출
    Dim p1 As Long, p2 As Long
    p1 = InStr(f, "[")
    p2 = InStr(p1, f, "]")
    If p1 > 0 And p2 > p1 Then
        ExtractExternalRef = Mid(f, p1 + 1, p2 - p1 - 1)
    End If
End Function

Private Function MaskFileOnly(fullRef As String) As String
    ' 파일명만 남기고 경로는 제거 (파일명 자체도 해시 마스킹은 안 함 — v1 기본)
    Dim slash As Long
    slash = InStrRev(fullRef, "\")
    If slash > 0 Then
        MaskFileOnly = Mid(fullRef, slash + 1)
    Else
        MaskFileOnly = fullRef
    End If
End Function

Private Function MaskPath(p As String) As String
    ' 사용자명 같은 민감 경로 마스킹
    Dim r As String: r = p
    r = Replace(r, Environ("USERPROFILE"), "~")
    r = Replace(r, Environ("USERNAME"), "<USER>")
    MaskPath = r
End Function

Private Function MaskExternal(ref As String) As String
    If InStr(ref, "[") > 0 And InStr(ref, ".xls") > 0 Then
        MaskExternal = "(external) " & MaskFileOnly(ref)
    Else
        MaskExternal = ref
    End If
End Function

Private Function Truncate(s As String, n As Long) As String
    If Len(s) > n Then
        Truncate = Left(s, n) & "..."
    Else
        Truncate = s
    End If
End Function

Private Function JoinFields(coll As Object) As String
    Dim pf As Object, out As String
    On Error Resume Next
    For Each pf In coll
        out = out & pf.Name & ", "
    Next pf
    On Error GoTo 0
    If Len(out) > 2 Then out = Left(out, Len(out) - 2)
    If Len(out) = 0 Then out = "(none)"
    JoinFields = out
End Function

Private Function FormatDictSorted(dict As Object, topN As Long) As String
    ' key-value 를 value 내림차순으로 정렬해 문자열화 (단순 선택 정렬)
    Dim keys() As String, vals() As Long, n As Long, i As Long, j As Long
    n = dict.Count
    If n = 0 Then
        FormatDictSorted = "    (none)" & vbCrLf
        Exit Function
    End If
    ReDim keys(n - 1)
    ReDim vals(n - 1)
    i = 0
    Dim k As Variant
    For Each k In dict.keys
        keys(i) = CStr(k)
        vals(i) = dict(k)
        i = i + 1
    Next k
    ' 선택 정렬
    For i = 0 To n - 2
        Dim maxI As Long: maxI = i
        For j = i + 1 To n - 1
            If vals(j) > vals(maxI) Then maxI = j
        Next j
        If maxI <> i Then
            Dim tv As Long: tv = vals(i): vals(i) = vals(maxI): vals(maxI) = tv
            Dim tk As String: tk = keys(i): keys(i) = keys(maxI): keys(maxI) = tk
        End If
    Next i
    Dim s As String, limit As Long
    limit = n - 1
    If topN > 0 And topN - 1 < limit Then limit = topN - 1
    For i = 0 To limit
        s = s & "    " & keys(i) & ": " & vals(i) & vbCrLf
    Next i
    FormatDictSorted = s
End Function

Private Function FormatPatternsTop(dict As Object, topN As Long) As String
    ' key = "sheet||blockTag||R1C1"  (blockTag 예: "B1", "B2" — SpecialCells Area 인덱스)
    ' value = "firstAddr|lastAddr|count|cseFlag"
    Dim n As Long, i As Long, j As Long
    n = dict.Count
    If n = 0 Then
        FormatPatternsTop = "    (none)" & vbCrLf
        Exit Function
    End If
    Dim keys() As String, counts() As Long, rngs() As String, cses() As String
    ReDim keys(n - 1): ReDim counts(n - 1): ReDim rngs(n - 1): ReDim cses(n - 1)
    i = 0
    Dim k As Variant, parts() As String
    For Each k In dict.keys
        keys(i) = CStr(k)
        parts = Split(dict(k), "|")
        ' parts(0)=first, parts(1)=last, parts(2)=count, parts(3)=cseFlag(optional)
        If parts(0) = parts(1) Then
            rngs(i) = parts(0)
        Else
            rngs(i) = parts(0) & ":" & parts(1)
        End If
        counts(i) = CLng(parts(2))
        If UBound(parts) >= 3 Then cses(i) = parts(3) Else cses(i) = "0"
        i = i + 1
    Next k
    ' 선택 정렬 (count desc)
    For i = 0 To n - 2
        Dim maxI As Long: maxI = i
        For j = i + 1 To n - 1
            If counts(j) > counts(maxI) Then maxI = j
        Next j
        If maxI <> i Then
            Dim tc As Long: tc = counts(i): counts(i) = counts(maxI): counts(maxI) = tc
            Dim tk As String: tk = keys(i): keys(i) = keys(maxI): keys(maxI) = tk
            Dim tr As String: tr = rngs(i): rngs(i) = rngs(maxI): rngs(maxI) = tr
            Dim tce As String: tce = cses(i): cses(i) = cses(maxI): cses(maxI) = tce
        End If
    Next i
    Dim s As String, limit As Long
    limit = n - 1
    If topN > 0 And topN - 1 < limit Then limit = topN - 1
    For i = 0 To limit
        Dim sheetName As String, blockTag As String, pattern As String, sp() As String
        sp = Split(keys(i), "||")
        sheetName = sp(0)
        If UBound(sp) >= 2 Then
            blockTag = sp(1)
            pattern = sp(2)
        Else
            blockTag = ""
            pattern = sp(1)
        End If
        s = s & "    [" & (i + 1) & "] " & sheetName
        If Len(blockTag) > 0 Then s = s & "!" & blockTag
        s = s & "!" & rngs(i) & " (" & counts(i) & " cells)"
        If cses(i) = "1" Then s = s & " [CSE array]"
        s = s & vbCrLf
        s = s & "        " & Truncate(pattern, 120) & vbCrLf
    Next i
    FormatPatternsTop = s
End Function

' sheet||block 별 집계 → 지배 패턴 >= 80% And 소수 패턴 == 1 cell 이면 아웃라이어 보고.
' 블록 총 셀수 < 5 인 경우는 신뢰도 낮아 skip.
Private Function FormatConsistencyAnomalies(patterns As Object) As String
    ' 전략: 1-cell 패턴마다, 같은 sheet+block 안에서 "동일한 col 을 사용하는"
    ' 다른 패턴 중 cnt>=5 인 것을 찾으면 outlier 로 본다. 블록 전체 80% 집계는
    ' 다컬럼 블록에서 dilute 되어 놓치는 케이스가 있어 pattern-pair 기반으로 변경.
    Dim s As String
    If patterns.Count = 0 Then
        FormatConsistencyAnomalies = "    (no formulas)" & vbCrLf
        Exit Function
    End If

    Dim reported As Long: reported = 0
    Dim k1 As Variant
    For Each k1 In patterns.keys
        Dim p1() As String
        p1 = Split(CStr(k1), "||")
        If UBound(p1) < 2 Then GoTo nxtOuter
        Dim v1() As String
        v1 = Split(patterns(k1), "|")
        If UBound(v1) < 2 Then GoTo nxtOuter
        If CLng(v1(2)) <> 1 Then GoTo nxtOuter  ' 1-cell 패턴만 후보

        Dim outAddr As String: outAddr = v1(0)
        Dim outCol As String: outCol = ColLettersOf(outAddr)
        Dim outBlock As String: outBlock = p1(0) & "||" & p1(1)

        Dim bestCnt As Long: bestCnt = 0
        Dim bestPat As String: bestPat = ""
        Dim k2 As Variant
        For Each k2 In patterns.keys
            If CStr(k1) = CStr(k2) Then GoTo nxtInner
            Dim p2() As String
            p2 = Split(CStr(k2), "||")
            If UBound(p2) < 2 Then GoTo nxtInner
            If (p2(0) & "||" & p2(1)) <> outBlock Then GoTo nxtInner
            Dim v2() As String
            v2 = Split(patterns(k2), "|")
            If UBound(v2) < 2 Then GoTo nxtInner
            Dim c2 As Long: c2 = CLng(v2(2))
            If c2 < 5 Then GoTo nxtInner
            ' 패턴 범위가 outlier 의 col 을 커버하는가?
            Dim firstCol As String: firstCol = ColLettersOf(v2(0))
            Dim lastCol As String: lastCol = ColLettersOf(v2(1))
            If firstCol = outCol And lastCol = outCol Then
                If c2 > bestCnt Then
                    bestCnt = c2
                    bestPat = p2(2)
                End If
            End If
nxtInner:
        Next k2

        If bestCnt >= 5 Then
            If reported >= 20 Then
                s = s & "    ... (more anomalies suppressed)" & vbCrLf
                FormatConsistencyAnomalies = s
                Exit Function
            End If
            s = s & "    - " & p1(0) & "!" & p1(1) & "!" & outAddr _
                & " outlier: " & Truncate(p1(2), 80) & vbCrLf
            s = s & "        dominant (" & bestCnt & " cells, same col): " & Truncate(bestPat, 80) & vbCrLf
            reported = reported + 1
        End If
nxtOuter:
    Next k1

    If reported = 0 Then s = "    (none - blocks look consistent)" & vbCrLf
    FormatConsistencyAnomalies = s
End Function

' "F250" -> "F", "AB100" -> "AB"
Private Function ColLettersOf(addr As String) As String
    Dim i As Long, c As String
    For i = 1 To Len(addr)
        c = Mid(addr, i, 1)
        If c >= "0" And c <= "9" Then Exit For
        ColLettersOf = ColLettersOf & c
    Next i
End Function

Private Function Has3DRef(f As String) As Boolean
    ' 3D ref pattern: Sheet1:Sheet2! (unquoted) or 'Sheet1:Sheet2'! (quoted).
    ' Uses VBScript.RegExp - avoids manual char scanning issues with CJK AscW.
    ' Excluded by char class: A1:B2 ranges (no trailing !),
    '   drive paths C:\, external refs '...\[file.xlsx]sheet'! (contains \ or [),
    '   structured refs [col]:[col] (contains []).
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = False
        re.IgnoreCase = False
        ' Unquoted: two tokens joined by ':' then '!'. Token excludes whitespace, punctuation that
        ' would terminate a sheet name, and path/bracket/quote/colon chars.
        ' [^ \t\r\n'()"\[\]\\/:!,;*?] - allow CJK (not excluded), allow letters/digits/_ /-
        re.Pattern = "[^ \t\r\n'()""\[\]\\/:!,;*?]+:[^ \t\r\n'()""\[\]\\/:!,;*?]+!"
    End If
    If re.Test(f) Then
        Has3DRef = True
        Exit Function
    End If
    ' Quoted variant: '...:...'!  where inside quotes has no \ or [ (excludes external refs)
    Static re2 As Object
    If re2 Is Nothing Then
        Set re2 = CreateObject("VBScript.RegExp")
        re2.Global = False
        re2.Pattern = "'[^'\\\[]+:[^'\\\[]+'!"
    End If
    Has3DRef = re2.Test(f)
End Function

Private Function CellAddr(ByVal r As Long, ByVal c As Long) As String
    ' 1-based row/col -> A1 주소 (Excel COM 호출 없이 계산)
    Dim col As String, n As Long, rmn As Long
    n = c
    Do While n > 0
        rmn = ((n - 1) Mod 26)
        col = Chr(65 + rmn) & col
        n = (n - 1) \ 26
    Loop
    CellAddr = col & r
End Function

Private Function IndentBlock(text As String, prefix As String) As String
    Dim lines() As String, i As Long, out As String
    lines = Split(text, vbLf)
    For i = LBound(lines) To UBound(lines)
        out = out & prefix & Replace(lines(i), vbCr, "") & vbCrLf
    Next i
    IndentBlock = out
End Function

'---------- Column Types ----------
' 헤더 추정:
'   1) ListObject 있으면 HeaderRowRange 우선
'   2) 상단 스캔 중 점수(str*3 - blank - num) 최대 행을 후보로
'   3) 후보 행에 number/date 셀이 있으면, 바로 다음 행과 타입 시그니처 비교 —
'      매칭 > 50% 이면 "헤더 없음" (headerRow=0, col1/col2 레이블, row 1 부터 샘플)
'   4) 점수가 임계치 미만이면 non-tabular
'   5) 후보+1 행이 후보와 비슷한 점수이면 multi-row header (rows=a-b)
Private Function SectionColumnTypes(wb As Workbook) As String
    Dim s As String, ws As Worksheet, u As Range, c As Range
    Dim nNum As Long, nStr As Long, nDate As Long, nFml As Long, nBlank As Long, nErr As Long
    s = "[Column Types]" & vbCrLf
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo nxtSheet
        On Error Resume Next
        Set u = ws.UsedRange
        On Error GoTo 0
        If u Is Nothing Then GoTo nxtSheet
        ' v1.2.1: 회사 실파일 적용 중 발견 — 임계값 30 은 너무 작아서
        ' step5 MB&camp map (127 cols), freshbag mapping (145 cols),
        ' step4 pivot (127/179 cols), 필요 쿼리 (38 cols) 등 PBI 분석에 핵심인
        ' 맵핑·피벗 시트가 전부 skip 되는 버그. 200 으로 상향 —
        ' step3 Summary New (1715 cols) 같은 wide 요약시트만 skip 되도록.
        If u.Columns.Count > 200 Or u.rows.Count < 2 Then
            s = s & "  - " & ws.Name & ": (skip, " & u.rows.Count & "x" & u.Columns.Count & ")" & vbCrLf
            GoTo nxtSheet
        End If

        Dim headerRow As Long, headerLastRow As Long, headerNote As String
        headerNote = ""
        GuessHeaderRange ws, u, headerRow, headerLastRow, headerNote

        Dim tag As String
        If headerRow = 0 Then
            tag = headerNote
            If Len(tag) = 0 Then tag = "no header detected"
        ElseIf headerLastRow > headerRow Then
            tag = "header rows=" & headerRow & "-" & headerLastRow
            If Len(headerNote) > 0 Then tag = tag & ", " & headerNote
        Else
            tag = "header row=" & headerRow
            If Len(headerNote) > 0 Then tag = tag & ", " & headerNote
        End If
        s = s & "  - " & ws.Name & ": (" & tag & ")" & vbCrLf

        Dim ci As Long, emptyCount As Long: emptyCount = 0
        Dim emptyLetters As String: emptyLetters = ""
        For ci = 1 To u.Columns.Count
            nNum = 0: nStr = 0: nDate = 0: nFml = 0: nBlank = 0: nErr = 0
            Dim ri As Long, sampleStart As Long, sampleLimit As Long
            If headerRow = 0 Then
                sampleStart = 1
            Else
                sampleStart = headerLastRow + 1
            End If
            sampleLimit = u.rows.Count
            If (sampleLimit - sampleStart + 1) > 200 Then sampleLimit = sampleStart + 199
            For ri = sampleStart To sampleLimit
                Set c = u.Cells(ri, ci)
                If IsError(c.Value) Then
                    nErr = nErr + 1
                ElseIf c.HasFormula Then
                    nFml = nFml + 1
                ElseIf IsEmpty(c.Value) Then
                    nBlank = nBlank + 1
                ElseIf IsDate(c.Value) Then
                    nDate = nDate + 1
                ElseIf IsNumeric(c.Value) Then
                    nNum = nNum + 1
                Else
                    nStr = nStr + 1
                End If
            Next ri

            If (nNum + nStr + nDate + nFml + nErr) = 0 Then
                emptyCount = emptyCount + 1
                If emptyCount <= 5 Then
                    Dim letter As String
                    letter = ColLettersOf(u.Cells(1, ci).Address(False, False))
                    If Len(emptyLetters) > 0 Then emptyLetters = emptyLetters & ","
                    emptyLetters = emptyLetters & letter
                End If
                GoTo nxtCol
            End If

            Dim header As String
            header = ""
            If headerRow = 0 Then
                header = "col" & ci
            ElseIf headerLastRow > headerRow Then
                ' 다단 헤더: row 1 셀이 실제 병합일 때만 상단+하단 결합.
                ' 병합 없이 header rows=1-2 로 검출된 경우는 1행이 진짜 헤더 +
                ' 2행은 데이터 첫 행일 가능성이 높으므로 row 1 값만 사용.
                Dim hTop As String, hBot As String, topMerged As Boolean
                On Error Resume Next
                topMerged = u.Cells(headerRow, ci).MergeCells
                If topMerged Then
                    hTop = CStr(u.Cells(headerRow, ci).MergeArea.Cells(1, 1).Value)
                Else
                    hTop = CStr(u.Cells(headerRow, ci).Value)
                End If
                hBot = CStr(u.Cells(headerLastRow, ci).Value)
                On Error GoTo 0
                If topMerged And Len(hTop) > 0 And Len(hBot) > 0 And hTop <> hBot Then
                    header = hTop & "/" & hBot
                ElseIf Len(hTop) > 0 Then
                    header = hTop
                ElseIf Len(hBot) > 0 Then
                    header = hBot
                End If
                If Len(header) = 0 Then header = "col" & ci
            Else
                On Error Resume Next
                header = CStr(u.Cells(headerRow, ci).Value)
                On Error GoTo 0
                If Len(header) = 0 Then header = "col" & ci
            End If

            Dim fmtSample As String, fmtMixed As Boolean, fmtDistinct As Long
            fmtSample = "": fmtMixed = False: fmtDistinct = 0
            Dim fmtSet As Object
            Set fmtSet = CreateObject("Scripting.Dictionary")
            Dim fmtRi As Long, curFmt As String
            For fmtRi = sampleStart To sampleLimit
                On Error Resume Next
                curFmt = CStr(u.Cells(fmtRi, ci).NumberFormat)
                On Error GoTo 0
                If Len(curFmt) > 0 And curFmt <> "General" Then
                    If Not fmtSet.Exists(curFmt) Then fmtSet.Add curFmt, 1
                    If Len(fmtSample) = 0 Then fmtSample = curFmt
                End If
            Next fmtRi
            fmtDistinct = fmtSet.Count
            If fmtDistinct >= 2 Then fmtMixed = True
            Set fmtSet = Nothing
            Dim hdrTag As String: hdrTag = ""
            If Len(header) > 0 And (Left(header, 1) = " " Or Right(header, 1) = " ") Then
                hdrTag = " [header has whitespace]"
            End If
            s = s & "      " & Truncate(header, 20) & " | num=" & nNum & " str=" & nStr _
                & " date=" & nDate & " fml=" & nFml & " blank=" & nBlank
            If nErr > 0 Then s = s & " err=" & nErr
            Dim totalData As Long: totalData = u.rows.Count - sampleStart + 1
            If (sampleLimit - sampleStart + 1) < totalData Then
                s = s & " (sampled " & (sampleLimit - sampleStart + 1) & "/" & totalData & ")"
            End If
            If Len(fmtSample) > 0 And fmtSample <> "General" Then
                s = s & " fmt=" & Truncate(fmtSample, 30)
                If fmtMixed Then s = s & " [mixed fmts: " & fmtDistinct & " distinct]"
            End If
            If fmtSample = "@" And nNum > 0 Then
                s = s & " [text-as-num suspect]"
            End If
            If fmtSample = "@" And nDate > 0 Then
                s = s & " [text-as-date suspect]"
            End If
            If Len(hdrTag) > 0 Then s = s & hdrTag
            s = s & vbCrLf
nxtCol:
        Next ci
        If emptyCount > 0 Then
            If emptyCount <= 5 Then
                s = s & "      (+" & emptyCount & " empty cols: " & emptyLetters & ")" & vbCrLf
            Else
                s = s & "      (+" & emptyCount & " empty cols, first 5: " & emptyLetters & ")" & vbCrLf
            End If
        End If
nxtSheet:
        Set u = Nothing
    Next ws
    SectionColumnTypes = s & vbCrLf
End Function

' 헤더 범위 추정 — byref 로 (firstRow, lastRow, note) 반환
' firstRow=0 → 헤더 없음 (row 1 부터 데이터)
Private Sub GuessHeaderRange(ws As Worksheet, u As Range, _
                              ByRef firstRow As Long, ByRef lastRow As Long, _
                              ByRef note As String)
    firstRow = 1: lastRow = 1: note = ""

    ' ListObject 우선
    Dim lo As Object
    On Error Resume Next
    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        If Not lo.HeaderRowRange Is Nothing Then
            firstRow = lo.HeaderRowRange.Row - u.Row + 1
            lastRow = firstRow
            If firstRow < 1 Then firstRow = 1: lastRow = 1
            Exit Sub
        End If
    End If
    On Error GoTo 0

    Dim scanMax As Long: scanMax = u.rows.Count
    If scanMax > 10 Then scanMax = 10

    Dim rowScore() As Long, rowStr() As Long, rowNum() As Long, rowBlank() As Long, rowFml() As Long
    ReDim rowScore(1 To scanMax)
    ReDim rowStr(1 To scanMax)
    ReDim rowNum(1 To scanMax)
    ReDim rowBlank(1 To scanMax)
    ReDim rowFml(1 To scanMax)

    Dim r As Long, ci As Long
    For r = 1 To scanMax
        For ci = 1 To u.Columns.Count
            Dim v As Variant
            Dim hf As Boolean: hf = False
            On Error Resume Next
            hf = u.Cells(r, ci).HasFormula
            v = u.Cells(r, ci).Value
            On Error GoTo 0
            If hf Then rowFml(r) = rowFml(r) + 1
            If IsEmpty(v) Then
                rowBlank(r) = rowBlank(r) + 1
            ElseIf IsError(v) Then
                ' 에러 반환 셀 — 헤더 판정에선 무시 (blank 처럼)
                rowBlank(r) = rowBlank(r) + 1
            ElseIf IsNumeric(v) Or IsDate(v) Then
                rowNum(r) = rowNum(r) + 1
            Else
                rowStr(r) = rowStr(r) + 1
            End If
        Next ci
        ' empty row disqualified from candidacy
        If rowStr(r) + rowNum(r) = 0 Then
            rowScore(r) = -999999
        Else
            rowScore(r) = rowStr(r) * 3 - rowBlank(r) - rowNum(r) * 2
        End If
    Next r

    Dim bestRow As Long: bestRow = 1
    Dim bestScore As Long: bestScore = rowScore(1)
    ' bias toward first row: strict > tiebreaker keeps earliest
    For r = 2 To scanMax
        If rowScore(r) > bestScore Then
            bestScore = rowScore(r)
            bestRow = r
        End If
    Next r

    ' 모든 후보가 음수이면 "ambiguous" — 첫 비어있지 않은 행을 보수적으로 선택하고 no-header 판정 금지
    Dim ambiguous As Boolean: ambiguous = False
    If bestScore < 0 Then
        ambiguous = True
        For r = 1 To scanMax
            If rowScore(r) > -999999 Then
                bestRow = r
                Exit For
            End If
        Next r
    End If

    ' pure numeric 전체: 모든 스캔 행에 string 셀이 하나도 없으면 no-header
    Dim totalStrAll As Long: totalStrAll = 0
    For r = 1 To scanMax
        totalStrAll = totalStrAll + rowStr(r)
    Next r
    If totalStrAll = 0 Then
        firstRow = 0
        lastRow = 0
        note = "no header (all numeric)"
        Exit Sub
    End If

    ' non-tabular 판단: 모든 행이 매우 낮은 score (str cell 1개 이하 시트 전체)
    Dim totalStr As Long: totalStr = 0
    For r = 1 To scanMax
        totalStr = totalStr + rowStr(r)
    Next r
    If totalStr <= 1 And u.Columns.Count <= 3 Then
        firstRow = 0
        note = "non-tabular"
        Exit Sub
    End If

    ' no-header 판단: ambiguous(전부 음수) 이면 스킵 (wide-pivot 등 오탐 방지)
    ' bestRow 에서 numeric cell 이 string cell 보다 많을 때만 시도 (YoY "Metric|2024|2025|Δ%" 처럼
    ' 연도 숫자가 헤더에 섞인 케이스 보호 — str>=num 이면 라벨 헤더로 신뢰)
    If Not ambiguous And bestRow < scanMax And rowNum(bestRow) > rowStr(bestRow) Then
        Dim matches As Long: matches = 0
        Dim total As Long: total = 0
        For ci = 1 To u.Columns.Count
            Dim v1 As Variant, v2 As Variant
            Dim t1 As Long, t2 As Long  ' 1=num/date, 2=str, 0=blank
            On Error Resume Next
            v1 = u.Cells(bestRow, ci).Value
            v2 = u.Cells(bestRow + 1, ci).Value
            On Error GoTo 0
            If IsEmpty(v1) Then
                t1 = 0
            ElseIf IsNumeric(v1) Or IsDate(v1) Then
                t1 = 1
            Else
                t1 = 2
            End If
            If IsEmpty(v2) Then
                t2 = 0
            ElseIf IsNumeric(v2) Or IsDate(v2) Then
                t2 = 1
            Else
                t2 = 2
            End If
            If t1 <> 0 And t2 <> 0 Then
                total = total + 1
                If t1 = t2 Then matches = matches + 1
            End If
        Next ci
        If total > 0 And matches * 2 >= total Then
            firstRow = 0
            note = "no header detected"
            Exit Sub
        End If
    End If

    firstRow = bestRow
    lastRow = bestRow
    If ambiguous Then note = "header ambiguous (wide-format?)"

    ' multi-row header: bestRow+1 도 str 위주이고 점수가 bestScore 의 60% 이상이면 편입.
    ' 단, (1) 수식 행은 데이터 (2) bestRow 에 실제 가로 병합이 있을 때만 확장.
    ' 병합 없는 pure-str 행은 대부분 데이터 (예: HiddenCols "V1-1/H1-1/V2-1/..." 샘플)
    If Not ambiguous And bestRow < scanMax Then
        Dim hasMerge As Boolean: hasMerge = False
        Dim mci As Long
        For mci = 1 To u.Columns.Count
            On Error Resume Next
            If u.Cells(bestRow, mci).MergeCells Then hasMerge = True
            On Error GoTo 0
            If hasMerge Then Exit For
        Next mci
        If hasMerge And rowStr(bestRow + 1) >= 2 And rowNum(bestRow + 1) = 0 And _
           rowFml(bestRow + 1) = 0 And _
           rowScore(bestRow + 1) * 10 >= bestScore * 6 Then
            lastRow = bestRow + 1
        End If
    End If
    ' merged-parent header: bestRow-1 이 병합된 상위 헤더일 가능성
    ' 조건 — bestRow >= 2 / bestRow-1 이 컨텐츠 있음(>=1) / 병합 slave 흔적(blank 비율 >= 25% 컬럼) /
    '        bestRow-1 이 full data row 아님 (non-blank < 컬럼 수)
    If Not ambiguous And bestRow >= 2 Then
        Dim prevNonBlank As Long
        prevNonBlank = rowStr(bestRow - 1) + rowNum(bestRow - 1)
        If prevNonBlank >= 1 And prevNonBlank < u.Columns.Count And _
           rowBlank(bestRow - 1) * 4 >= u.Columns.Count Then
            firstRow = bestRow - 1
        End If
    End If
End Sub

'---------- Merged Cells ----------
' 스케일 전략:
'   1) ws.UsedRange.MergeCells 한 번만 질의 → True/False/Null 중 하나.
'   2) False = 병합 없음 (가장 흔함). Null = 일부 병합, True = 전부 병합 (드묾).
'   3) 병합이 있는 경우 상단 30행 × 전체 열만 셀 순회 — 재무 파일 병합은 사실상 헤더 한정.
Private Function SectionMergedCells(wb As Workbook) As String
    Dim s As String, ws As Worksheet, total As Long
    s = "[Merged Cells]" & vbCrLf
    For Each ws In wb.Worksheets
        Dim mc As Variant
        mc = Null
        On Error Resume Next
        mc = ws.UsedRange.MergeCells
        On Error GoTo 0

        ' False 면 완전히 없음 — 스킵
        If Not IsNull(mc) Then
            If mc = False Then GoTo nxtMC
        End If

        Dim scanRows As Long, scanCols As Long
        scanRows = ws.UsedRange.rows.Count
        If scanRows > 30 Then scanRows = 30
        scanCols = ws.UsedRange.Columns.Count

        Dim probe As Range, r As Range, cnt As Long
        On Error Resume Next
        Set probe = ws.UsedRange.Resize(scanRows, scanCols)
        On Error GoTo 0
        If probe Is Nothing Then GoTo nxtMC

        cnt = 0
        On Error Resume Next
        For Each r In probe
            If r.MergeCells Then
                If r.Address = r.MergeArea.Cells(1, 1).Address Then
                    cnt = cnt + 1
                    If cnt <= 5 Then
                        s = s & "  - " & ws.Name & "!" & r.MergeArea.Address(False, False) & vbCrLf
                    End If
                End If
            End If
        Next r
        On Error GoTo 0
        If cnt > 5 Then s = s & "    ... (+" & (cnt - 5) & " more in " & ws.Name _
            & ", header-only scan first " & scanRows & " rows)" & vbCrLf
        total = total + cnt
nxtMC:
        Set probe = Nothing
    Next ws
    If total = 0 Then s = s & "  (none in top 30 header-scan rows)" & vbCrLf
    SectionMergedCells = s & vbCrLf
End Function

'---------- Comments / Notes (legacy + threaded, dedup by cell) ----------
' Threaded 가 있으면 legacy 는 skip (threaded 가 최신 진실).
Private Function SectionComments(wb As Workbook) As String
    Dim s As String, ws As Worksheet, cm As Comment, cnt As Long
    s = "[Comments / Notes]" & vbCrLf
    For Each ws In wb.Worksheets
        Dim threadedKeys As Object
        Set threadedKeys = CreateObject("Scripting.Dictionary")

        ' Threaded comments (modern Excel) 먼저 — 우선권
        On Error Resume Next
        Dim ct As Object
        For Each ct In ws.CommentsThreaded
            cnt = cnt + 1
            Dim addrKey As String: addrKey = ""
            addrKey = ct.Parent.Address(False, False)
            If Len(addrKey) > 0 Then threadedKeys(addrKey) = True
            Dim tText As String, tAuthor As String
            tText = "": tAuthor = ""
            tText = ct.text
            tAuthor = ct.Author.Name
            s = s & "  - [threaded] " & ws.Name & "!" & addrKey _
                & " (" & tAuthor & "): " & Truncate(Replace(tText, vbLf, " "), 100) & vbCrLf
            Set ct = Nothing
        Next ct
        On Error GoTo 0

        ' Legacy notes — threaded 에 없는 셀만
        On Error Resume Next
        For Each cm In ws.Comments
            Dim lAddr As String: lAddr = ""
            lAddr = cm.Parent.Address(False, False)
            If Not threadedKeys.Exists(lAddr) Then
                cnt = cnt + 1
                Dim txt As String
                txt = ""
                txt = cm.text
                s = s & "  - [note] " & ws.Name & "!" & lAddr & ": " _
                    & Truncate(Replace(txt, vbLf, " "), 100) & vbCrLf
            End If
        Next cm
        On Error GoTo 0

        Set threadedKeys = Nothing
    Next ws
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionComments = s & vbCrLf
End Function

'---------- Hyperlinks ----------
Private Function SectionHyperlinks(wb As Workbook) As String
    Dim s As String, ws As Worksheet, hl As Hyperlink, cnt As Long
    s = "[Hyperlinks]" & vbCrLf
    For Each ws In wb.Worksheets
        On Error Resume Next
        For Each hl In ws.Hyperlinks
            cnt = cnt + 1
            Dim target As String, kind As String
            If Len(hl.Address) > 0 Then
                Dim lowAddr As String: lowAddr = LCase$(hl.Address)
                If InStr(lowAddr, "http://") = 1 Or InStr(lowAddr, "https://") = 1 Then
                    kind = "[web]"
                    target = hl.Address
                ElseIf InStr(lowAddr, "mailto:") = 1 Then
                    kind = "[mail]"
                    target = hl.Address
                ElseIf InStr(lowAddr, "file://") = 1 Or InStr(hl.Address, "\") > 0 Or InStr(hl.Address, ":/") > 0 Then
                    kind = "[file]"
                    target = MaskPath(hl.Address)
                Else
                    kind = "[ext]"
                    target = MaskPath(hl.Address)
                End If
                If Len(hl.SubAddress) > 0 Then target = target & "#" & hl.SubAddress
            Else
                kind = "[internal]"
                target = hl.SubAddress
            End If
            ' v1.2.1: 하이퍼링크 URL 은 절대 truncate 하지 않음.
            ' Google Sheets 링크(100자+)가 `...`로 잘려서 스켈레톤 재구성 때 링크가 깨지던 문제.
            s = s & "  - " & kind & " " & ws.Name & "!" & hl.Range.Address(False, False) _
                & " -> " & target & vbCrLf
        Next hl
        On Error GoTo 0
    Next ws
    If cnt = 0 Then s = s & "  (none)" & vbCrLf
    SectionHyperlinks = s & vbCrLf
End Function

'---------- AutoFilters ----------
Private Function SectionAutoFilters(wb As Workbook) As String
    Dim s As String, ws As Worksheet, cnt As Long, i As Long
    s = "[AutoFilters]" & vbCrLf
    For Each ws In wb.Worksheets
        ' Sheet-level AutoFilter
        On Error Resume Next
        If ws.AutoFilterMode Then
            Dim af As Object
            Set af = ws.AutoFilter
            If Not af Is Nothing Then
                cnt = cnt + 1
                s = s & "  - " & ws.Name & "  range=" & af.Range.Address(False, False) _
                    & "  fields=" & af.Filters.Count & vbCrLf
                For i = 1 To af.Filters.Count
                    If af.Filters(i).On Then
                        Dim crit1 As String: crit1 = ""
                        crit1 = CStr(af.Filters(i).Criteria1)
                        s = s & "      field " & i & ": " & Truncate(crit1, 60) & vbCrLf
                    End If
                Next i
            End If
            Set af = Nothing
        End If

        ' ListObject-level AutoFilter
        Dim lo As ListObject
        For Each lo In ws.ListObjects
            If Not lo.AutoFilter Is Nothing Then
                Dim anyActive As Boolean: anyActive = False
                For i = 1 To lo.AutoFilter.Filters.Count
                    If lo.AutoFilter.Filters(i).On Then anyActive = True
                Next i
                If anyActive Then
                    cnt = cnt + 1
                    s = s & "  - " & ws.Name & "!" & lo.Name & " (table)  fields=" _
                        & lo.AutoFilter.Filters.Count & vbCrLf
                    For i = 1 To lo.AutoFilter.Filters.Count
                        If lo.AutoFilter.Filters(i).On Then
                            Dim lc1 As String: lc1 = ""
                            lc1 = CStr(lo.AutoFilter.Filters(i).Criteria1)
                            s = s & "      field " & i & ": " & Truncate(lc1, 60) & vbCrLf
                        End If
                    Next i
                End If
            End If
        Next lo
        On Error GoTo 0
    Next ws
    If cnt = 0 Then s = s & "  (none active)" & vbCrLf
    SectionAutoFilters = s & vbCrLf
End Function

'---------- Shapes (TextBoxes, Buttons, OLEObjects, Pictures) ----------
' 모든 shape property 접근을 On Error Resume Next 로 감싼다 — 중첩 핸들러 금지.
Private Function SectionShapes(wb As Workbook) As String
    On Error GoTo localErr
    Dim s As String, total As Long
    s = "[Shapes]" & vbCrLf
    Dim i As Long, j As Long
    For i = 1 To wb.Worksheets.Count
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(i)
        On Error GoTo localErr
        If ws Is Nothing Then GoTo nxtWs
        LogSection "Shapes." & ws.Name, "iter"
        Dim shCount As Long: shCount = 0
        On Error Resume Next
        shCount = ws.Shapes.Count
        On Error GoTo localErr
        If shCount = 0 Then GoTo nxtWs
        Dim sCount As Long: sCount = 0
        Dim buf As String: buf = ""
        For j = 1 To shCount
            Dim shpName As String: shpName = ""
            Dim shpType As Long: shpType = -1
            Dim extra As String: extra = ""
            On Error Resume Next
            shpName = ws.Shapes(j).Name
            shpType = ws.Shapes(j).Type
            If shpType <> 3 And shpType <> 25 And shpType <> -2 Then
                Dim t As String: t = ""
                Dim hasTF As Long: hasTF = 0
                hasTF = ws.Shapes(j).HasTextFrame
                If hasTF <> 0 Then
                    t = ws.Shapes(j).TextFrame2.TextRange.text
                    If Len(t) > 0 Then extra = " text=""" & Truncate(Replace(t, vbLf, " "), 40) & """"
                End If
                Dim ona As String: ona = ""
                ona = ws.Shapes(j).OnAction
                If Len(ona) > 0 Then extra = extra & " onAction=" & ona
            End If
            On Error GoTo localErr
            If shpType = 3 Then GoTo nxtShape  ' Chart
            If shpType = 4 Then GoTo nxtShape  ' Comment — [Comments/Notes] 섹션이 별도 커버
            sCount = sCount + 1
            buf = buf & "      " & shpName & "  [" & ShapeTypeName(shpType) & "]" & extra & vbCrLf
nxtShape:
        Next j
        If sCount > 0 Then
            s = s & "  - " & ws.Name & ": " & sCount & vbCrLf & buf
            total = total + sCount
        End If
nxtWs:
    Next i
    If total = 0 Then s = s & "  (none - charts listed separately)" & vbCrLf
    SectionShapes = s & vbCrLf
    Exit Function
localErr:
    SectionShapes = s & "  (shapes section error: " & Err.Number & " " & Err.Description & ")" & vbCrLf & vbCrLf
End Function

Private Function ShapeTypeName(t As Long) As String
    Select Case t
        Case 1: ShapeTypeName = "AutoShape"
        Case 3: ShapeTypeName = "Chart"
        Case 4: ShapeTypeName = "Comment"
        Case 6: ShapeTypeName = "Group"
        Case 7: ShapeTypeName = "EmbeddedOLE"
        Case 8: ShapeTypeName = "FormControl"
        Case 9: ShapeTypeName = "Line"
        Case 10: ShapeTypeName = "LinkedOLEObject"
        Case 11: ShapeTypeName = "LinkedPicture"
        Case 12: ShapeTypeName = "ActiveXControl"
        Case 13: ShapeTypeName = "Picture"
        Case 14: ShapeTypeName = "Placeholder"
        Case 17: ShapeTypeName = "TextBox"
        Case 19: ShapeTypeName = "Callout"
        Case 23: ShapeTypeName = "Diagram"
        Case 24: ShapeTypeName = "OLEControl"
        Case 25: ShapeTypeName = "Slicer"
        Case 27: ShapeTypeName = "WebVideo"
        Case 28: ShapeTypeName = "ContentApp"
        Case -2: ShapeTypeName = "Timeline"
        Case Else: ShapeTypeName = "type" & t
    End Select
End Function

'---------- Scenarios (What-If Manager) ----------
Private Function SectionScenarios(wb As Workbook) As String
    On Error GoTo localErr
    Dim s As String, total As Long
    s = "[Scenarios]" & vbCrLf
    Dim i As Long, j As Long, k As Long
    For i = 1 To wb.Worksheets.Count
        Dim ws As Worksheet
        Set ws = wb.Worksheets(i)
        Dim scCount As Long: scCount = 0
        On Error Resume Next
        scCount = ws.Scenarios.Count
        On Error GoTo localErr
        If scCount = 0 Then GoTo nxtScWs
        For j = 1 To scCount
            Dim scName As String: scName = ""
            Dim scAddr As String: scAddr = ""
            Dim vals As Variant
            Dim scComment As String: scComment = ""
            On Error Resume Next
            scName = ws.Scenarios(j).Name
            scAddr = ws.Scenarios(j).ChangingCells.Address(False, False)
            vals = ws.Scenarios(j).Values
            scComment = ws.Scenarios(j).Comment
            On Error GoTo localErr
            Dim valStr As String: valStr = ""
            If IsArray(vals) Then
                For k = LBound(vals) To UBound(vals)
                    If Len(valStr) > 0 Then valStr = valStr & ", "
                    valStr = valStr & CStr(vals(k))
                    If k - LBound(vals) >= 5 Then
                        valStr = valStr & ", ..."
                        Exit For
                    End If
                Next k
            End If
            total = total + 1
            s = s & "  - " & ws.Name & "!" & scName & _
                "  changing=" & scAddr & _
                "  values=(" & valStr & ")"
            If Len(scComment) > 0 Then
                s = s & "  [" & Truncate(scComment, 30) & "]"
            End If
            s = s & vbCrLf
        Next j
nxtScWs:
    Next i
    If total = 0 Then s = s & "  (none)" & vbCrLf
    SectionScenarios = s & vbCrLf
    Exit Function
localErr:
    SectionScenarios = s & "  (scenarios section error: " & Err.Number & " " & Err.Description & ")" & vbCrLf & vbCrLf
End Function

'---------- Data Model (Power Pivot) ----------
Private Function SectionDataModel(wb As Workbook) As String
    Dim s As String
    s = "[Data Model]" & vbCrLf
    On Error Resume Next
    Dim mdl As Object
    Set mdl = wb.Model
    If mdl Is Nothing Then
        s = s & "  (no data model)" & vbCrLf & vbCrLf
        SectionDataModel = s
        Exit Function
    End If

    Dim tCount As Long, mCount As Long, rCount As Long
    tCount = 0: mCount = 0: rCount = 0
    tCount = mdl.ModelTables.Count
    mCount = mdl.ModelMeasures.Count
    rCount = mdl.ModelRelationships.Count

    If tCount = 0 And mCount = 0 Then
        s = s & "  (no data model)" & vbCrLf & vbCrLf
        SectionDataModel = s
        Exit Function
    End If

    s = s & "  Tables: " & tCount & "  Measures: " & mCount & "  Relationships: " & rCount & vbCrLf

    Dim mt As Object, mi As Long
    mi = 0
    For Each mt In mdl.ModelTables
        mi = mi + 1
        If mi > 20 Then
            s = s & "    ... (+" & (tCount - 20) & " more tables)" & vbCrLf
            Exit For
        End If
        s = s & "    - table: " & mt.Name
        On Error Resume Next
        s = s & "  records=" & mt.RecordCount
        On Error GoTo 0
        s = s & vbCrLf
    Next mt

    Dim mm As Object, mj As Long
    mj = 0
    For Each mm In mdl.ModelMeasures
        mj = mj + 1
        If mj > 30 Then
            s = s & "    ... (+" & (mCount - 30) & " more measures)" & vbCrLf
            Exit For
        End If
        Dim tblName As String: tblName = ""
        On Error Resume Next
        tblName = mm.AssociatedTable.Name
        On Error GoTo 0
        s = s & "    - measure: " & mm.Name & " @ " & tblName _
            & "  formula=" & Truncate(CStr(mm.Formula), 80) & vbCrLf
    Next mm
    On Error GoTo 0
    SectionDataModel = s & vbCrLf
End Function

'---------- Sparklines ----------
Private Function SectionSparklines(wb As Workbook) As String
    Dim s As String, ws As Worksheet, total As Long
    s = "[Sparklines]" & vbCrLf
    For Each ws In wb.Worksheets
        On Error Resume Next
        Dim u As Range
        Set u = ws.UsedRange
        If u Is Nothing Then GoTo nxtSP
        Dim sg As Object, sgCount As Long: sgCount = 0
        sgCount = u.SparklineGroups.Count
        If sgCount = 0 Then GoTo nxtSP
        s = s & "  - " & ws.Name & ": " & sgCount & " group(s)" & vbCrLf
        Dim i As Long
        For i = 1 To sgCount
            Set sg = u.SparklineGroups.Item(i)
            Dim tName As String
            Select Case sg.Type
                Case 1: tName = "Line"
                Case 2: tName = "Column"
                Case 3: tName = "Win/Loss"
                Case Else: tName = "type" & sg.Type
            End Select
            Dim loc As String, src As String
            loc = "": src = ""
            loc = sg.Location.Address(False, False)
            src = sg.SourceData
            s = s & "      [" & i & "] " & tName & " loc=" & loc _
                & " src=" & Truncate(src, 50) & vbCrLf
            Set sg = Nothing
        Next i
        total = total + sgCount
nxtSP:
        Set u = Nothing
        On Error GoTo 0
    Next ws
    If total = 0 Then s = s & "  (none)" & vbCrLf
    SectionSparklines = s & vbCrLf
End Function

'---------- View (Frozen Panes + Print Area + Print Titles) ----------
' 주의: FreezePanes 상태는 Window 레벨이라 시트 활성화가 필요하다.
' DispatchEx 로 열린 숨김 인스턴스라 활성화는 저비용. 다만 에러 발생 시 조용히 스킵.
' View section — Print_Area와 FreezePanes를 [Named Ranges]에서 이미 덮으므로
' 별도 섹션으로 만들지 않고 [Named Ranges] 에서 Print_Area/Print_Titles 가 자동 노출된다.
' FreezePanes 는 v1.3 에서 재설계 (현재는 수집기 hang 이슈로 skip).

'---------- Grouping / Outline ----------
' v1.1: skipped.
'   이유 — Range.OutlineLevel 은 속성 질의 시 Excel 이 Range 내 모든 행/열을 내부적으로 평가한다.
'   Worksheet.Rows (1,048,576 행 전체) 는 물론, UsedRange.Rows 여도 수만줄 스케일에서는 30초+ 멈춤이 재현됐다.
'   재무 파일에서 그룹화 감지는 중요도가 낮아 v1 에서는 (skipped) 로 표기하고 v2 에서 대안 탐색.
Private Function SectionGrouping(wb As Workbook) As String
    SectionGrouping = "[Grouping / Outline]" & vbCrLf _
        & "  (skipped in v1 - Range.OutlineLevel hangs at large scale)" & vbCrLf & vbCrLf
End Function

'---------- Output ----------
Private Sub WriteToNewWorkbook(text As String, srcName As String)
    Dim out As Workbook, ws As Worksheet
    Set out = Workbooks.Add
    Set ws = out.Sheets(1)
    ws.Name = "Report"
    ws.Range("A1").Value = "Source: " & srcName
    ws.Range("A2").Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A3").Value = "--- REPORT BELOW (A4) ---"

    ' 개별 라인으로 A4 부터 주입 (Excel 셀 문자 한계 32767 회피)
    Dim lines() As String, i As Long
    lines = Split(text, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        ws.Cells(4 + i, 1).Value = "'" & lines(i)   ' 앞에 ' 붙여서 수식 해석 방지
    Next i
    ws.Columns("A").ColumnWidth = 120
End Sub
