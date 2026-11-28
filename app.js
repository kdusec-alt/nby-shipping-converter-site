(() => {
  const state = { files: [], sheets: [], result: null, activeTab: 'orders', templates: {}, busy: false, replacingKind: '', statusMessage: null, sourceOverrides: {} };
  const $ = id => document.getElementById(id);
  let language='zh';
  const i18n={
    zh:{bulkTitle:'一次匯入全部來源檔',bulkHint:'一次選取本次 Nike FC、SAP ZRSD1019、SAP MB52、UPS、SF Baozun 與 AWB 檔案；系統會依欄位自動辨識並分開驗證。人工成品參考檔仍可在個別區匯入。',bulkButton:'選取全部來源檔',categoryUpload:'個別補檔或替換來源（選填）',optionalUpload:'選填參考檔：AWB 或人工成品',bulkProcessing:'正在處理第 {current} / {total} 個檔案…',bulkImported:'已檢查 {count} 個檔案；成功辨識 {recognized} 個來源檔。請確認來源清單與日期。',brandSub:'每日資料核對與轉檔 · v2',privacy:'檔案僅在此瀏覽器處理',eyebrow:'每日出貨資料工作區',heroTitle:'出貨資料轉檔與逐筆核對',lede:'匯入 Nike FC、SAP、UPS 與 SF Baozun 檔案，再與人工報表逐筆比對。檔案僅在瀏覽器本機處理。',dateLabel:'本次出貨基準日',uploadTitle:'1. 分類匯入來源檔',uploadHint:'選取全部來源檔一次匯入，系統會依欄位自動辨識並分別驗證；不同日期或重複批次不會自動合併。',addFiles:'新增檔案',replaceFile:'替換檔案',noFiles:'尚未匯入檔案',addAwb:'新增 AWB',replaceAwb:'替換 AWB',manualReference:'人工成品／Shipping Status 歷史（選填）',chooseReference:'選擇參考檔',dateOverrides:'進階日期設定（僅有不同報表日期時修改）',dateOverrideHint:'預設使用上方的本次出貨基準日。只有報表日期不同時才需要填寫；留空即沿用基準日。',ordersDate:'Orders Update 日期',dataDate:'Data Shipping 日期',statusStart:'Shipping Status 起日',statusEnd:'Shipping Status 迄日',fcSourceLabel:'本次 FC 來源',importFirst:'請先匯入來源',templateTitle:'正式 Excel 樣板（保留原始格式、公式與工作表）',templateHint:'分別載入既有 .xlsx 成品作為樣板。樣板不作為原始資料；未提供時輸出 DRAFT 資料版。若樣板結構無法安全填入，該正式檔不會產生。',ordersTemplate:'Orders Update 樣板',dataTemplate:'Data Shipping 樣板',statusTemplate:'Shipping Status 樣板',validationHeading:'驗收分層',validationNote:'檔案讀取、必要欄位驗證、來源匹配、報表產生與人工成品驗收分開顯示。出貨統計目前是工單數；Golden Sample 尚未認證。',run:'執行轉檔與核對',clear:'清除本次資料',resultsEyebrow:'出貨資料核對結果',resultsTitle:'2. 核對與轉檔結果',downloadZip:'下載本次報表 ZIP',tabOrders:'工單關聯',tabUps:'UPS 獨立核對',tabSf:'SF Baozun 核對',tabExceptions:'異常清單',tabManual:'人工成品驗證',search:'搜尋工單、Tracking、SAP SD Document…',footnote:'轉檔輸出與核對結果皆由本機匯入資料即時計算。缺少必要來源時會標示「待來源」，不會推測補值。',footer:'內部測試工具 · NBY Digitalization · 不連接正式系統',switchLanguage:'切換語言',readFile:'正在讀取檔案…',validateColumns:'正在驗證欄位…',importDone:'匯入完成。請檢查資料類型與欄位後再執行。',importFailed:'匯入失敗',missingRequired:'缺少必要欄位',importSuccess:'匯入成功／必要欄位已驗證',recommendedMissing:'已匯入並通過必要欄位檢查；建議欄位缺漏：{columns}',readFailed:'讀取失敗',columnWarning:'欄位警告',parsed:'已解析',rowsColumns:'{rows} 列 · {columns} 欄 · {range}',preview:'預覽前 10 列',notDetected:'未偵測',revalidate:'重新驗證',remove:'移除',waitingFile:'等待檔案',missingSource:'缺少來源',optionalMissing:'選填／未匯入',multipleBatches:'同類來源有多個活頁簿，請確認下方來源選擇',ready:'已就緒／必要欄位已驗證',warning:'警告／請檢查欄位',runProcessing:'正在匹配來源並套用 Excel 樣板…',runDone:'處理流程完成；請查看各報表狀態。待驗證不等於已驗收。',runFailed:'核對失敗',dataChanged:'資料或日期已變更，請重新執行',templateLoaded:'已載入樣板：{name}',zipUnavailable:'Excel 已產生；ZIP 函式庫載入失敗，請個別下載。',outputNote:'Expect Shipment 未計算；尚未完成正式驗收。',rowsCount:'{count} 筆',rowsLimited:'{count} 筆（畫面顯示前 500 筆）',metricOrders:'Orders Update FC 列',metricData:'Data Shipping FC 列',metricStatus:'Shipping Status FC 列',metricSap:'SAP ZRSD1019',metricMb:'MB52 庫存列',metricExceptions:'待處理異常',statusUnit:'工單數不等同鞋雙數',statusSap:'工單關聯來源',statusMb:'快照，不代表出貨量',statusRules:'依來源與關聯規則產生',sourceNote:'採用 FC 工作表',sourceMatch:'{count} 筆與 SAP 關聯 {rate}%',sourceNoteText:'三份輸出各用獨立日期範圍；匹配率僅代表工單關聯，不代表數量正確。',tableRows:'{count} 筆',tableRowsLimited:'{count} 筆（畫面顯示前 500 筆）',pass:'通過',yes:'是',no:'否',reportMissing:'缺少必要來源',reportGenerated:'已產生／待驗證',reportDraft:'已產生草稿／待驗證',reportFailed:'產生失敗',reportReview:'已產生／待人工檢視',reportGolden:'待提供獨立原始檔驗證',reportOrders:'Orders Update',reportData:'Data Shipping',reportStatus:'Shipping Status',reportUps:'UPS 核對報表',reportSf:'SF Baozun 核對報表',reportException:'異常報告',reportGoldenDiff:'Golden Sample 差異報告',reportGoldenCert:'Golden Sample 驗收',genericSource:'來源檔',noSheet:'沒有可重新驗證的工作表',validDate:'請選擇有效工單核對基準日',statusDateOrder:'Shipping Status 起日不可晚於迄日',fcMissing:'請先選擇本次 FC 來源工作表',noFcRows:'所選 FC 來源沒有符合核對／報表日期的資料',invalidDate:'請設定有效的日期',fcRequired:'Nike FC 缺少必要欄位',multipleSource:'有多個來源工作表，請移除非本次批次後再執行',multipleSourceBooks:'有多個來源檔案；請保留本次要核對的單一批次。',tableSearchCount:'筆',selectSource:'請選擇 FC 工作表',sourceRecords:'列',sourceColumns:'欄',sourceDateRange:'至',download:'下載',wrongSource:'此區需上傳 {expected}；辨識結果：{actual}',unrecognizedSheet:'找不到可辨識的工作表',allImportsFailed:'所有檔案皆匯入失敗，請檢查檔案錯誤訊息'},
    en:{bulkTitle:'Import All Source Files at Once',bulkHint:'Select the Nike FC, SAP ZRSD1019, SAP MB52, UPS, SF Baozun, and AWB files for this run in one file picker. The app identifies each file from its columns and validates sources separately. Manual reference reports can still be added below.',bulkButton:'Select All Source Files',categoryUpload:'Add or Replace Individual Sources (Optional)',optionalUpload:'Optional references: AWB or manual reports',bulkProcessing:'Processing file {current} of {total}…',bulkImported:'Checked {count} files and recognized {recognized} source files. Review the source list and dates.',brandSub:'Daily Source Reconciliation · v2',privacy:'Files stay in this browser',eyebrow:'DAILY SHIPPING WORKSPACE',heroTitle:'Shipping Data Conversion and Row-by-Row Reconciliation',lede:'Import Nike FC, SAP, UPS, and SF Baozun files, then compare each row with the manually prepared reports. Files are processed locally in your browser.',dateLabel:'Shipment Date for This Run',uploadTitle:'1. Upload Source Files by Category',uploadHint:'Select all source files for one run and import them together. The app identifies each file from its columns and validates it separately; different dates and duplicate batches are never merged automatically.',addFiles:'Add Files',replaceFile:'Replace Files',noFiles:'No files imported',addAwb:'Add AWB Mapping',replaceAwb:'Replace AWB Mapping',manualReference:'Manual Reports / Shipping Status History (Optional)',chooseReference:'Choose Reference Files',dateOverrides:'Advanced Date Overrides (only if report dates differ)',dateOverrideHint:'The date above is used for all reports by default. Enter an override only when a report uses a different date or period; leave blank to use the date above.',ordersDate:'Orders Update Date',dataDate:'Data Shipping Date',statusStart:'Shipping Status Start Date',statusEnd:'Shipping Status End Date',fcSourceLabel:'FC Source for This Run',importFirst:'Import a source file first',templateTitle:'Official Excel Templates (Preserve Original Formatting, Formulas, and Worksheets)',templateHint:'Load each existing .xlsx report as its template. Templates are not source data. Without a template, the app creates a clearly labeled DRAFT data workbook. If a template cannot be filled safely, that official workbook will not be generated.',ordersTemplate:'Orders Update Template',dataTemplate:'Data Shipping Template',statusTemplate:'Shipping Status Template',validationHeading:'Validation Statuses Are Separate',validationNote:'File reading, required-column validation, source matching, report generation, and manual-report verification are shown separately. Shipment totals currently count work orders, not confirmed shoe pairs. The Golden Sample has not been certified.',run:'Generate Reports and Reconcile',clear:'Clear This Run',resultsEyebrow:'SHIPPING DATA RECONCILIATION',resultsTitle:'2. Reconciliation and Reports',downloadZip:'Download All Reports (ZIP)',tabOrders:'Work Order Matching',tabUps:'UPS Reconciliation',tabSf:'SF Baozun Reconciliation',tabExceptions:'Exceptions',tabManual:'Manual Report Validation',search:'Search work order, tracking number, or SAP SD document…',footnote:'Reports and reconciliation results are calculated locally from the imported files. Missing required sources are flagged; the app does not guess missing values.',footer:'Internal test tool · NBY Digitalization · No connection to production systems',switchLanguage:'Switch language',readFile:'Reading file…',validateColumns:'Validating columns…',importDone:'Import complete. Review the detected file type and columns before generating reports.',importFailed:'Import failed',missingRequired:'Missing required columns',importSuccess:'Import successful / required columns validated',recommendedMissing:'Imported and required columns validated; recommended columns missing: {columns}',readFailed:'Read failed',columnWarning:'Column warning',parsed:'Parsed',rowsColumns:'{rows} rows · {columns} columns · {range}',preview:'Preview first 10 rows',notDetected:'Not detected',revalidate:'Revalidate',remove:'Remove',waitingFile:'Waiting for file',missingSource:'Missing source',optionalMissing:'Optional / not provided',multipleBatches:'Multiple workbooks for this source; confirm the selection below',ready:'Ready / required columns validated',warning:'Warning / review columns',runProcessing:'Matching sources and applying Excel templates…',runDone:'Processing finished. Review each report status; pending validation does not mean verified.',runFailed:'Reconciliation failed',dataChanged:'Source data or dates changed. Run reconciliation again.',templateLoaded:'Template loaded: {name}',zipUnavailable:'Excel files were generated, but the ZIP library failed to load. Download each file individually.',outputNote:'Expect Shipment was not calculated. Formal validation is still pending.',rowsCount:'{count} rows',rowsLimited:'{count} rows (showing the first 500)',metricOrders:'Orders Update FC rows',metricData:'Data Shipping FC rows',metricStatus:'Shipping Status FC rows',metricSap:'SAP ZRSD1019',metricMb:'MB52 stock rows',metricExceptions:'Open exceptions',statusUnit:'Work order count; not confirmed shoe-pair quantity',statusSap:'Order-matching source',statusMb:'Inventory snapshot; not shipment quantity',statusRules:'Generated from source and matching rules',sourceNote:'FC worksheet used',sourceMatch:'{count} work orders; {rate}% matched to SAP',sourceNoteText:'Each report uses its own date filter. Match rate reflects work-order linkage only, not quantity or report accuracy.',tableRows:'{count} rows',tableRowsLimited:'{count} rows (showing first 500)',pass:'Matched',yes:'Yes',no:'No',reportMissing:'Missing Required Source',reportGenerated:'Generated · Pending Validation',reportDraft:'Draft Generated · Pending Validation',reportFailed:'Failed',reportReview:'Generated · Pending Review',reportGolden:'Pending independent raw-source validation',reportOrders:'Orders Update',reportData:'Data Shipping',reportStatus:'Shipping Status',reportUps:'UPS Reconciliation',reportSf:'SF Baozun Reconciliation',reportException:'Exception Report',reportGoldenDiff:'Golden Sample Difference Report',reportGoldenCert:'Golden Sample Certification',genericSource:'Source file',noSheet:'No worksheet available to revalidate',validDate:'Select a valid work-order reconciliation date',statusDateOrder:'Shipping Status start date must be on or before its end date',fcMissing:'Select the FC worksheet for this run',noFcRows:'The selected FC source has no rows for the chosen reconciliation/report dates',invalidDate:'Set a valid date for',fcRequired:'Nike FC is missing required columns',multipleSource:'contains multiple source worksheets.',multipleSourceBooks:'has multiple source workbooks. Keep only the workbook for this run.',tableSearchCount:'rows',selectSource:'Choose an FC worksheet',sourceRecords:'rows',sourceColumns:'columns',sourceDateRange:'to',download:'Download',wrongSource:'This area requires {expected}; detected: {actual}',unrecognizedSheet:'No recognizable worksheet found',allImportsFailed:'All files failed to import. Review the file error messages.'},
    vi:{bulkTitle:'Tải tất cả tệp nguồn cùng lúc',bulkHint:'Chọn các tệp Nike FC, SAP ZRSD1019, SAP MB52, UPS, SF Baozun và AWB của lần chạy này trong cùng một cửa sổ. Hệ thống tự nhận diện theo cột và xác thực riêng từng nguồn. Có thể tải thêm báo cáo thủ công ở phần bên dưới.',bulkButton:'Chọn tất cả tệp nguồn',categoryUpload:'Bổ sung hoặc thay thế từng nguồn (Tùy chọn)',optionalUpload:'Tệp tham chiếu tùy chọn: AWB hoặc báo cáo thủ công',bulkProcessing:'Đang xử lý tệp {current}/{total}…',bulkImported:'Đã kiểm tra {count} tệp và nhận diện được {recognized} tệp nguồn. Vui lòng kiểm tra danh sách nguồn và ngày.',brandSub:'Đối chiếu dữ liệu hằng ngày · v2',privacy:'Tệp được xử lý ngay trong trình duyệt',eyebrow:'KHÔNG GIAN XỬ LÝ DỮ LIỆU GIAO HÀNG',heroTitle:'Chuyển đổi dữ liệu giao hàng và đối chiếu từng dòng',lede:'Tải tệp Nike FC, SAP, UPS và SF Baozun lên, sau đó đối chiếu từng dòng với các báo cáo đã lập thủ công. Tệp được xử lý cục bộ trong trình duyệt.',dateLabel:'Ngày giao hàng của lần xử lý này',uploadTitle:'1. Tải tệp nguồn theo danh mục',uploadHint:'Chọn tất cả tệp nguồn của một lần chạy và tải lên cùng lúc. Hệ thống nhận diện theo cột và xác thực riêng từng tệp; không tự động gộp các ngày hoặc lô tải trùng nhau.',addFiles:'Thêm tệp',replaceFile:'Thay thế tệp',noFiles:'Chưa có tệp nào được tải lên',addAwb:'Thêm bảng AWB',replaceAwb:'Thay bảng AWB',manualReference:'Báo cáo thủ công / Lịch sử Shipping Status (Tùy chọn)',chooseReference:'Chọn tệp tham chiếu',dateOverrides:'Cài đặt ngày nâng cao (chỉ chỉnh khi ngày báo cáo khác nhau)',dateOverrideHint:'Mặc định dùng ngày giao hàng ở trên cho tất cả báo cáo. Chỉ nhập ngày riêng khi báo cáo cần ngày hoặc khoảng thời gian khác; để trống để dùng ngày phía trên.',ordersDate:'Ngày Orders Update',dataDate:'Ngày Data Shipping',statusStart:'Ngày bắt đầu Shipping Status',statusEnd:'Ngày kết thúc Shipping Status',fcSourceLabel:'Nguồn FC cho lần chạy này',importFirst:'Vui lòng tải tệp nguồn lên trước',templateTitle:'Mẫu Excel chính thức (Giữ nguyên định dạng, công thức và trang tính)',templateHint:'Tải từng báo cáo .xlsx hiện có làm mẫu. Mẫu không phải là dữ liệu nguồn. Nếu không có mẫu, hệ thống tạo tệp dữ liệu DRAFT có nhãn rõ ràng. Nếu không thể điền mẫu an toàn, báo cáo chính thức đó sẽ không được tạo.',ordersTemplate:'Mẫu Orders Update',dataTemplate:'Mẫu Data Shipping',statusTemplate:'Mẫu Shipping Status',validationHeading:'Trạng thái xác thực được hiển thị riêng',validationNote:'Đọc tệp, kiểm tra cột bắt buộc, đối chiếu nguồn, tạo báo cáo và xác minh với báo cáo thủ công được hiển thị thành các bước riêng. Tổng giao hàng hiện đếm theo Work Order, chưa xác nhận tương đương số đôi giày. Golden Sample chưa được chứng nhận.',run:'Tạo báo cáo và đối chiếu',clear:'Xóa dữ liệu lần chạy này',resultsEyebrow:'KẾT QUẢ ĐỐI CHIẾU DỮ LIỆU GIAO HÀNG',resultsTitle:'2. Kết quả đối chiếu và báo cáo',downloadZip:'Tải tất cả báo cáo (ZIP)',tabOrders:'Đối chiếu Work Order',tabUps:'Đối chiếu UPS',tabSf:'Đối chiếu SF Baozun',tabExceptions:'Danh sách ngoại lệ',tabManual:'Xác minh báo cáo thủ công',search:'Tìm Work Order, số Tracking hoặc SAP SD Document…',footnote:'Báo cáo và kết quả đối chiếu được tính cục bộ từ các tệp đã tải lên. Nguồn bắt buộc còn thiếu sẽ được cảnh báo; hệ thống không tự suy đoán dữ liệu.',footer:'Công cụ thử nghiệm nội bộ · NBY Digitalization · Không kết nối hệ thống sản xuất',switchLanguage:'Đổi ngôn ngữ',readFile:'Đang đọc tệp…',validateColumns:'Đang kiểm tra cột…',importDone:'Đã tải tệp. Vui lòng kiểm tra loại tệp và các cột được nhận diện trước khi tạo báo cáo.',importFailed:'Tải tệp thất bại',missingRequired:'Thiếu cột bắt buộc',importSuccess:'Tải thành công / đã xác thực cột bắt buộc',recommendedMissing:'Đã tải và xác thực cột bắt buộc; còn thiếu cột khuyến nghị: {columns}',readFailed:'Đọc tệp thất bại',columnWarning:'Cảnh báo về cột',parsed:'Đã phân tích',rowsColumns:'{rows} dòng · {columns} cột · {range}',preview:'Xem trước 10 dòng đầu',notDetected:'Không phát hiện',revalidate:'Kiểm tra lại',remove:'Xóa',waitingFile:'Đang chờ tệp',missingSource:'Thiếu nguồn',optionalMissing:'Tùy chọn / chưa cung cấp',multipleBatches:'Có nhiều sổ làm việc cho nguồn này; hãy xác nhận lựa chọn bên dưới',ready:'Sẵn sàng / đã xác thực cột bắt buộc',warning:'Cảnh báo / vui lòng kiểm tra cột',runProcessing:'Đang đối chiếu nguồn và áp dụng mẫu Excel…',runDone:'Đã xử lý xong. Vui lòng xem trạng thái từng báo cáo; chờ xác thực không đồng nghĩa đã xác minh.',runFailed:'Đối chiếu thất bại',dataChanged:'Dữ liệu nguồn hoặc ngày đã thay đổi. Vui lòng chạy đối chiếu lại.',templateLoaded:'Đã tải mẫu: {name}',zipUnavailable:'Đã tạo các tệp Excel nhưng thư viện ZIP không tải được. Vui lòng tải từng tệp riêng.',outputNote:'Chưa tính Expect Shipment. Việc xác thực chính thức vẫn đang chờ.',rowsCount:'{count} dòng',rowsLimited:'{count} dòng (đang hiển thị 500 dòng đầu)',metricOrders:'Số dòng FC cho Orders Update',metricData:'Số dòng FC cho Data Shipping',metricStatus:'Số dòng FC cho Shipping Status',metricSap:'ZRSD1019 SAP',metricMb:'Số dòng tồn kho MB52',metricExceptions:'Ngoại lệ cần xử lý',statusUnit:'Đếm Work Order; chưa xác nhận là số đôi giày',statusSap:'Nguồn đối chiếu đơn hàng',statusMb:'Ảnh chụp tồn kho; không phải lượng giao hàng',statusRules:'Được tạo theo dữ liệu nguồn và quy tắc đối chiếu',sourceNote:'Trang tính FC được sử dụng',sourceMatch:'{count} Work Order; {rate}% đã khớp SAP',sourceNoteText:'Mỗi báo cáo dùng bộ lọc ngày riêng. Tỷ lệ khớp chỉ phản ánh liên kết Work Order, không xác nhận số lượng hoặc độ chính xác của báo cáo.',tableRows:'{count} dòng',tableRowsLimited:'{count} dòng (đang hiển thị 500 dòng đầu)',pass:'Khớp',yes:'Có',no:'Không',reportMissing:'Thiếu nguồn bắt buộc',reportGenerated:'Đã tạo · Chờ xác thực',reportDraft:'Đã tạo bản nháp · Chờ xác thực',reportFailed:'Thất bại',reportReview:'Đã tạo · Chờ xem xét',reportGolden:'Chờ xác minh bằng dữ liệu nguồn độc lập',reportOrders:'Orders Update',reportData:'Data Shipping',reportStatus:'Shipping Status',reportUps:'Báo cáo đối chiếu UPS',reportSf:'Báo cáo đối chiếu SF Baozun',reportException:'Báo cáo ngoại lệ',reportGoldenDiff:'Báo cáo chênh lệch Golden Sample',reportGoldenCert:'Xác nhận Golden Sample',genericSource:'Tệp nguồn',noSheet:'Không có trang tính để kiểm tra lại',validDate:'Vui lòng chọn ngày đối chiếu Work Order hợp lệ',statusDateOrder:'Ngày bắt đầu Shipping Status phải trước hoặc bằng ngày kết thúc',fcMissing:'Vui lòng chọn trang tính FC cho lần chạy này',noFcRows:'Nguồn FC đã chọn không có dòng dữ liệu trong các ngày đối chiếu/báo cáo đã chọn',invalidDate:'Vui lòng chọn ngày hợp lệ cho',fcRequired:'Nike FC thiếu cột bắt buộc',multipleSource:'có nhiều trang tính nguồn.',multipleSourceBooks:'có nhiều tệp nguồn. Chỉ giữ lại tệp thuộc lô cần xử lý.',tableSearchCount:'dòng',selectSource:'Chọn trang tính FC',sourceRecords:'dòng',sourceColumns:'cột',sourceDateRange:'đến',download:'Tải xuống',wrongSource:'Khu vực này cần tệp {expected}; nhận diện được: {actual}',unrecognizedSheet:'Không tìm thấy trang tính có thể nhận diện',allImportsFailed:'Tất cả tệp đều tải lên thất bại. Vui lòng xem thông báo lỗi.'}
  };
  Object.assign(i18n.zh,{sourceOverridesTitle:'同類來源有多份檔案，請指定本次要用的檔案',sourceOverridesHint:'系統不會自動合併不同活頁簿。檔名或工作表名稱明確時會先建議來源；其他檔案仍保留在匯入清單。',chooseSourceBook:'偵測到多個 {kind} 來源，請在下方選擇本次要使用的檔案。',sourceAutoSelected:'依檔名或工作表名稱建議',sourceManualSelected:'已由你指定本次來源',sourceSelectionLabel:'{kind} 來源檔'});
  Object.assign(i18n.en,{sourceOverridesTitle:'Multiple files were detected for the same source. Choose the file for this run.',sourceOverridesHint:'Workbooks are not merged automatically. A clear filename or worksheet name is suggested when available; all files remain in the import list.',chooseSourceBook:'Multiple {kind} workbooks were detected. Choose the file to use below.',sourceAutoSelected:'Suggested from the filename or worksheet name',sourceManualSelected:'Selected for this run',sourceSelectionLabel:'{kind} source file'});
  Object.assign(i18n.vi,{sourceOverridesTitle:'Có nhiều tệp cho cùng một nguồn. Hãy chọn tệp dùng cho lần xử lý này.',sourceOverridesHint:'Hệ thống không tự động gộp các sổ làm việc. Nếu tên tệp hoặc trang tính rõ ràng, hệ thống sẽ đề xuất nguồn; tất cả tệp vẫn được giữ trong danh sách tải lên.',chooseSourceBook:'Phát hiện nhiều tệp {kind}. Vui lòng chọn tệp dùng bên dưới.',sourceAutoSelected:'Được đề xuất theo tên tệp hoặc trang tính',sourceManualSelected:'Đã chọn cho lần xử lý này',sourceSelectionLabel:'Tệp nguồn {kind}'});
  const phraseMaps={
    zh:{},
    en:{'通過':'Matched','有差異':'Difference','高':'High','中':'Medium','低':'Low','有':'Present','未填':'Blank','等待檔案':'Waiting for file','缺少來源':'Missing source','匯入完成。請檢查資料類型與欄位後再執行核對。':'Import complete. Review source types and columns before reconciliation.','待確認':'Pending','參考相同':'Same (reference only)','參考不同':'Different (reference only)','是':'Yes','否':'No','SAP 訂單匹配':'SAP Order Matched','MB52 庫存紀錄匹配':'MB52 Stock Record Matched','UPS Tracking 匹配':'UPS Tracking Matched','AWB 狀態':'AWB Status','已':'Yes','FC 工單缺少 SAP ZRSD1019':'FC Work Order Missing from SAP ZRSD1019','SAP 訂單缺少 MB52 庫存紀錄':'SAP Order Missing from MB52 Stock Snapshot','FC Tracking 未出現在 UPS':'FC Tracking Number Missing from UPS','FC 缺少 Tracking Number':'FC Work Order Missing Tracking Number','未找到 AWB／Consolidation ID':'AWB / Consolidation ID Not Found','UPS Tracking 重複':'Duplicate UPS Tracking Number','UPS Tracking 未關聯 FC 工單':'UPS Tracking Not Linked to an FC Work Order','SF manifest 工單未匹配 FC':'SF Work Order Not Matched to FC','FC 工單重複':'Duplicate FC Work Order','FC 缺少 Work Order Nbr':'FC Work Order Number Missing','FC 工單明細粒度待確認':'FC Work Order Detail Grain Needs Review','SAP 關聯不唯一':'SAP Match Is Ambiguous','AWB 對照衝突':'Conflicting AWB Mapping','未辨識物流公司':'Carrier Not Recognized','FC Tracking 重複':'Duplicate FC Tracking Number','UPS 缺少 Tracking Number':'UPS Tracking Number Missing','來源檔讀取失敗或未辨識':'Source File Read Failed or Unrecognized','來源缺少必要欄位':'Required Source Columns Missing','來源未提供或未通過欄位驗證':'Source Missing or Failed Column Validation','AWB Mapping 未提供':'AWB Mapping Not Provided','下載批次待確認':'Download Batch Needs Review','人工成品逐筆比對差異':'Manual Report Row-by-Row Difference','資料或日期已變更，請重新執行':'Source data or dates changed. Run reconciliation again.','沒有可辨識的資料工作表':'No recognizable data worksheet found','原始來源無此列':'Row Not Found in Source','欄位相符':'Fields Matched','工單欄位':'Work Order Fields','UPS Tracking':'UPS Tracking','SF Baozun manifest':'SF Baozun Manifest','MB52 庫存列':'MB52 Stock Row','工單逐筆／欄位':'Work Order Row / Fields','UPS Tracking 逐筆':'UPS Tracking Rows','SF Baozun 逐筆':'SF Baozun Rows','MB52 庫存列逐筆':'MB52 Stock Rows','尚未匯入人工成品':'No Manual Reports Imported','工單及來源已匹配（非數量核對）':'Work order and source matched (not a quantity check)','Tracking 不一致':'Tracking Number Mismatch','Ship Group 不一致':'Ship Group Mismatch','FC Tracking 一致；數量未認證':'FC Tracking Matches; Quantity Not Certified','FC Tracking 一致；SAP Confirmed Qty 僅供參考，MB52 數量待判讀':'FC Tracking Matches; SAP Confirmed Qty Is for Reference; MB52 Needs Business Review','SF 工單未出現在 FC':'SF Work Order Not Found in FC','SF 多筆明細，數量粒度待確認':'Multiple SF Detail Rows; Quantity Grain Needs Review','SF Qty 與 SAP Confirmed Qty 參考差異（待判讀）':'SF Qty Differs from SAP Confirmed Qty (Reference Only; Review Pending)','MB52 庫存快照差額（待業務判讀）':'MB52 Stock Snapshot Delta (Business Review Pending)','不等同當日出貨量':'Not equivalent to daily shipment quantity','工單數不等同鞋雙數':'Work order count is not shoe-pair quantity','快照，不代表出貨量':'Snapshot; not shipment quantity','工單關聯來源':'Work-order matching source','依來源與關聯規則產生':'Based on source and matching rules','所選 FC 來源沒有符合核對／報表日期的資料':'Selected FC source has no data for the selected reconciliation/report dates','請先選擇本次 FC 來源工作表':'Select an FC worksheet for this run','請選擇有效工單核對基準日':'Select a valid work-order reconciliation date','Shipping Status 起日不可晚於迄日':'Shipping Status start date cannot be after its end date','核對失敗':'Reconciliation failed'},
    vi:{'通過':'Khớp','有差異':'Có chênh lệch','高':'Cao','中':'Trung bình','低':'Thấp','有':'Có dữ liệu','未填':'Để trống','等待檔案':'Đang chờ tệp','缺少來源':'Thiếu nguồn','匯入完成。請檢查資料類型與欄位後再執行核對。':'Đã tải tệp. Vui lòng kiểm tra loại nguồn và cột trước khi đối chiếu.','待確認':'Đang chờ xác nhận','參考相同':'Trùng khớp (chỉ tham khảo)','參考不同':'Khác biệt (chỉ tham khảo)','是':'Có','否':'Không','SAP 訂單匹配':'SAP đã khớp đơn hàng','MB52 庫存紀錄匹配':'Đã khớp bản ghi tồn kho MB52','UPS Tracking 匹配':'Đã khớp Tracking UPS','AWB 狀態':'Trạng thái AWB','FC 工單缺少 SAP ZRSD1019':'Work Order FC không có trong SAP ZRSD1019','SAP 訂單缺少 MB52 庫存紀錄':'Đơn SAP không có bản ghi tồn kho MB52','FC Tracking 未出現在 UPS':'Tracking FC không có trong dữ liệu UPS','FC 缺少 Tracking Number':'Work Order FC thiếu Tracking Number','未找到 AWB／Consolidation ID':'Không tìm thấy AWB / Consolidation ID','UPS Tracking 重複':'Tracking UPS bị trùng','UPS Tracking 未關聯 FC 工單':'Tracking UPS chưa liên kết với Work Order FC','SF manifest 工單未匹配 FC':'Work Order SF chưa khớp với FC','FC 工單重複':'Work Order FC bị trùng','FC 缺少 Work Order Nbr':'Thiếu Work Order Nbr trong FC','FC 工單明細粒度待確認':'Cần xem lại mức độ chi tiết Work Order FC','SAP 關聯不唯一':'Liên kết SAP không duy nhất','AWB 對照衝突':'Ánh xạ AWB bị xung đột','未辨識物流公司':'Không nhận diện được đơn vị vận chuyển','FC Tracking 重複':'Tracking FC bị trùng','UPS 缺少 Tracking Number':'UPS thiếu Tracking Number','來源檔讀取失敗或未辨識':'Không đọc được hoặc không nhận diện tệp nguồn','來源缺少必要欄位':'Tệp nguồn thiếu cột bắt buộc','來源未提供或未通過欄位驗證':'Thiếu nguồn hoặc không đạt xác thực cột','AWB Mapping 未提供':'Chưa cung cấp bảng ánh xạ AWB','下載批次待確認':'Cần xác nhận lô tải xuống','人工成品逐筆比對差異':'Chênh lệch khi đối chiếu từng dòng báo cáo thủ công','原始來源無此列':'Không có dòng này trong nguồn','欄位相符':'Các trường khớp','工單欄位':'Các trường Work Order','UPS Tracking':'Tracking UPS','SF Baozun manifest':'Bảng kê SF Baozun','MB52 庫存列':'Dòng tồn kho MB52','工單逐筆／欄位':'Đối chiếu Work Order / trường dữ liệu','UPS Tracking 逐筆':'Đối chiếu từng Tracking UPS','SF Baozun 逐筆':'Đối chiếu từng dòng SF Baozun','MB52 庫存列逐筆':'Đối chiếu từng dòng tồn kho MB52','尚未匯入人工成品':'Chưa tải báo cáo thủ công lên','工單及來源已匹配（非數量核對）':'Work Order và nguồn đã khớp (không phải đối chiếu số lượng)','Tracking 不一致':'Tracking không khớp','Ship Group 不一致':'Ship Group không khớp','FC Tracking 一致；數量未認證':'Tracking FC khớp; số lượng chưa được chứng nhận','FC Tracking 一致；SAP Confirmed Qty 僅供參考，MB52 數量待判讀':'Tracking FC khớp; SAP Confirmed Qty chỉ để tham khảo; cần xem xét MB52','SF 工單未出現在 FC':'Không tìm thấy Work Order SF trong FC','SF 多筆明細，數量粒度待確認':'Nhiều dòng chi tiết SF; cần xác nhận mức độ chi tiết số lượng','SF Qty 與 SAP Confirmed Qty 參考差異（待判讀）':'SF Qty khác SAP Confirmed Qty (chỉ tham khảo; chờ xem xét)','MB52 庫存快照差額（待業務判讀）':'Chênh lệch ảnh chụp tồn kho MB52 (chờ nghiệp vụ xem xét)','不等同當日出貨量':'Không tương đương số lượng giao trong ngày','工單數不等同鞋雙數':'Số Work Order không đồng nghĩa số đôi giày','快照，不代表出貨量':'Ảnh chụp tồn kho; không phải lượng giao hàng','工單關聯來源':'Nguồn liên kết Work Order','依來源與關聯規則產生':'Dựa trên quy tắc nguồn và liên kết','所選 FC 來源沒有符合核對／報表日期的資料':'Nguồn FC đã chọn không có dữ liệu trong ngày đối chiếu/báo cáo đã chọn','請先選擇本次 FC 來源工作表':'Vui lòng chọn trang tính FC cho lần chạy này','請選擇有效工單核對基準日':'Vui lòng chọn ngày đối chiếu Work Order hợp lệ','Shipping Status 起日不可晚於迄日':'Ngày bắt đầu Shipping Status không được sau ngày kết thúc','核對失敗':'Đối chiếu thất bại'}
  };
  const languageOrder=['zh','en','vi'];
  const t=(key,params={})=>String(i18n[language]?.[key]??i18n.zh[key]??key).replace(/\{(\w+)\}/g,(_,name)=>params[name]??'');
  const uiText=text=>phraseMaps[language]?.[String(text??'')]??String(text??'');
  const status=(key,params={})=>{state.statusMessage={key,params};$('runStatus').textContent=t(key,params);};
  const norm = v => String(v ?? '').trim().replace(/\.0$/, '').replace(/\s+/g, ' ').toUpperCase();
  const value = (row, ...keys) => {
    for (const key of keys) { const k = Object.keys(row || {}).find(x => norm(x) === norm(key)); if (k && row[k] !== '' && row[k] != null) return row[k]; }
    return '';
  };
  const cleanDate = v => {
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
    const s = String(v ?? '').trim(); if (!s) return '';
    let m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    m=s.match(/^(\d{4})\.(\d{1,2})\.(\d{1,2})$/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    m=s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);if(m){const month=new Date(`${m[2]} 1, 2000`).getMonth()+1;return `${m[3].length===2?'20'+m[3]:m[3]}-${String(month).padStart(2,'0')}-${m[1].padStart(2,'0')}`;}
    m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);if(m)return `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
    const d = new Date(s); if (!Number.isNaN(d.valueOf())) return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    return s;
  };
  window.NBYCleanDate=cleanDate;
  const dateForInput = () => $('asOfDate').value;
  const keyBy = (rows, fn) => { const m = new Map(); rows.forEach(r => { const k = norm(fn(r)); if (!k) return; if (!m.has(k)) m.set(k, []); m.get(k).push(r); }); return m; };
  const kindOf = (rows, fileName='') => {
    const first = rows?.[0] || {}; const keys = Object.keys(first).map(norm); const has = (...k) => k.every(x => keys.includes(norm(x)));
    if (has('Date','Actually shipment','Expect shipment')) return 'Shipping Status Summary';
    if (has('SAP 訂單匹配') || has('核對狀態') || has('Work Order Nbr','SAP SD Document ZRSD')) return '人工成品';
    if (has('Work Order Nbr','Manifest Scan Date') && (has('ZRSD1019') || has('MB52') || has('Pick up date') || has('Report date'))) return '人工成品';
    if (has('Work Order Nbr','Manifest Scan Date','Shipper')) return 'Nike FC';
    if (has('OrderNumber','ConsolidationShipmentId')) return 'AWB Map';
    if (has('Work Order Number','Tracking Number')) return 'SF Baozun';
    if (has('Tracking Number','Ship Date') && (has('Service') || /UPS|Detailed Report/i.test(fileName))) return 'UPS';
    if (has('Customer Reference','SD Document')) return 'SAP ZRSD1019';
    if (has('SD Document','Material') && (has('Unrestricted') || has('Blocked Stock'))) return 'SAP MB52';
    return '';
  };
  function rowsFromSheet(sheet, expectedKind="") {
    const grid = XLSX.utils.sheet_to_json(sheet, {header:1, defval:'', raw:false, blankrows:false});
    let headerAt = -1;
    for (let i=0;i<Math.min(grid.length,15);i++) {
      const line = grid[i].map(norm);
      if ((line.includes('WORK ORDER NBR') && line.includes('MANIFEST SCAN DATE')) ||
          (expectedKind==='Nike FC' && (line.includes('WORK ORDER NBR') || line.includes('MANIFEST SCAN DATE'))) ||
          (expectedKind==='SAP ZRSD1019' && (line.includes('CUSTOMER REFERENCE') || line.includes('SD DOCUMENT'))) ||
          (expectedKind==='SAP MB52' && (line.includes('SD DOCUMENT') || line.includes('MATERIAL'))) ||
          (expectedKind==='UPS' && line.includes('TRACKING NUMBER')) ||
          (expectedKind==='SF Baozun' && (line.includes('WORK ORDER NUMBER') || line.includes('TRACKING NUMBER') || line.includes('MOTHER WAYBILL'))) ||
          (line.includes('CUSTOMER REFERENCE') && line.includes('SD DOCUMENT')) ||
          (line.includes('PLANT') && line.includes('SD DOCUMENT') && line.includes('MATERIAL')) ||
          (line.includes('TRACKING NUMBER') && (line.includes('SHIP DATE') || line.includes('MOTHER WAYBILL'))) ||
          (line.includes('TRACKING NUMBER') && (line.includes('UPS 原始出現次數') || line.includes('紀錄次數') || line.includes('重複筆數'))) ||
          (line.includes('WORK ORDER NUMBER') && line.includes('TRACKING NUMBER')) ||
          (line.includes('DATE') && line.includes('ACTUALLY SHIPMENT') && line.includes('EXPECT SHIPMENT')) ||
          (line.includes('ORDERNUMBER') && line.includes('CONSOLIDATIONSHIPMENTID')) || line.includes('SAP 訂單匹配')) { headerAt=i; break; }
    }
    if (headerAt < 0) return [];
    const seen={}; const headers = grid[headerAt].map((h,i)=>{const k=String(h||`欄位${i+1}`).trim();const n=seen[k]||0;seen[k]=n+1;return n?`${k}.${n}`:k;});
    return grid.slice(headerAt+1).filter(r=>r.some(x=>String(x).trim()!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
  }
  function statusDetailRows(sheet) {
    const grid=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:false,blankrows:false});
    const headerAt=grid.findIndex(row=>row.some(x=>norm(x)==='PRODUCT NAME'));
    if(headerAt<0)return [];
    const headers=grid[headerAt].map((h,i)=>String(h||`Column ${i+1}`).trim());
    return grid.slice(headerAt+1).filter(row=>row.some(x=>String(x).trim()!=='')).map(row=>Object.fromEntries(headers.map((h,i)=>[h,row[i]??''])));
  }

  async function parseFile(file, expectedKind="") {
    const data = new Uint8Array(await file.arrayBuffer());
    const wb = XLSX.read(data, {type:'array', cellDates:true});
    const manualBook = expectedKind==='人工成品' || expectedKind==='REFERENCE' || /NBY_1120|NBY_Shipping|Source_Analysis|完整源頭|人工成品/i.test(file.name);
    const statusBook = expectedKind==='Shipping Status' || expectedKind==='REFERENCE' || /SHI+PPING\s+STATUS/i.test(file.name) || (wb.SheetNames.some(n=>/^summary$/i.test(n)) && wb.SheetNames.some(n=>/^detail shipment$/i.test(n)));
    const parsed=[];
    for (const name of wb.SheetNames) {
      if(statusBook && /^summary$/i.test(name)){const rows=rowsFromSheet(wb.Sheets[name]);if(rows.length)parsed.push({file:file.name,sheet:name,kind:'Shipping Status Summary',rows});continue;}
      if(statusBook && /^detail shipment$/i.test(name)){const rows=statusDetailRows(wb.Sheets[name]);if(rows.length)parsed.push({file:file.name,sheet:name,kind:'Shipping Status Detail',rows});continue;}
      const rows = rowsFromSheet(wb.Sheets[name],expectedKind); if (!rows.length) continue;
      const detectedKind=kindOf(rows,file.name); const kind=detectedKind==='人工成品'?'人工成品':expectedKind==='REFERENCE'?'人工成品':(expectedKind?expectedKind:(detectedKind|| (manualBook?'人工成品':''))); if (kind) parsed.push({file:file.name,sheet:name,kind,rows,processed:rows.some(r=>Object.hasOwn(r,'ZRSD1019')||Object.hasOwn(r,'MB52'))});
    }
    return parsed;
  }
  async function ingest(fileList, expectedKind="") {
    const incoming=[...fileList]; if (!incoming.length || state.busy) return; invalidate();state.busy=true;let recognized=0;$('runButton').disabled=true;
    status('readFile');
    for (const [fileIndex,file] of incoming.entries()) {
      if(!expectedKind)status('bulkProcessing',{current:fileIndex+1,total:incoming.length});
      if(state.files.some(f=>f.name===file.name)){state.files=state.files.filter(f=>f.name!==file.name);state.sheets=state.sheets.filter(s=>s.file!==file.name);}
      try { let parsed=await parseFile(file,expectedKind);if(expectedKind&&expectedKind!=='REFERENCE'){const accepted=parsed.filter(x=>x.kind===expectedKind);if(!accepted.length)throw Error(t('wrongSource',{expected:expectedKind,actual:[...new Set(parsed.map(x=>x.kind))].join(', ')||t('unrecognizedSheet')}));parsed=accepted;}if(!parsed.length)throw Error(t('unrecognizedSheet'));status('validateColumns');await new Promise(resolve=>setTimeout(resolve,0));const validation=validateFileRows(parsed);state.files.push({name:file.name,size:file.size,importedAt:new Date().toISOString(),expectedKind,types:[...new Set(parsed.map(x=>x.kind))],sheetCount:parsed.length,validation,sourceFile:file}); state.sheets.push(...parsed);recognized++;
        const candidate=parsed.find(x=>x.kind==='人工成品'&&x.rows.some(r=>Object.hasOwn(r,'ZRSD1019')&&Object.hasOwn(r,'MB52')));
        if(candidate&&!state.templates.data){
          const message={zh:'偵測到 Data Shipping 人工成品。是否將此 XLSX 設為固定樣板並儲存在此瀏覽器？',en:'A Data Shipping manual report was detected. Save this XLSX as the template in this browser?',vi:'Đã nhận diện báo cáo Data Shipping. Lưu tệp XLSX này làm mẫu trong trình duyệt?'}[language];
          if(!/\.xlsx$/i.test(file.name)){
            status('xlsxTemplateRequired');
          }else if(window.confirm(message)){
            try{const template={bytes:new Uint8Array(await file.arrayBuffer()),name:file.name};saveTemplate('data',template);state.templates.data=template;renderTemplateNames();}
            catch(error){console.warn(error);status('templateSaveFailed');}
          }
        } }
      catch(e) { state.files.push({name:file.name,size:file.size,importedAt:new Date().toISOString(),expectedKind,types:['讀取失敗'],sheetCount:0,error:e.message,sourceFile:file}); }
    }
    state.busy=false;renderSources(); renderSourceChoices();renderSourceOverrides();status(state.files.some(f=>!f.error&&f.sheetCount)?(expectedKind?'importDone':'bulkImported'):'allImportsFailed',expectedKind?{}:{count:incoming.length,recognized}); $('runButton').disabled=state.sheets.length===0; $('clearButton').disabled=false;
  }
  const schemas={
    'Nike FC':{required:['Work Order Nbr','Manifest Scan Date','Shipper'],recommended:['Product Name','Tracking Number','Ship Group Nbr']},
    'SAP ZRSD1019':{required:['Customer Reference','SD Document'],recommended:['Material',['Item (SD)','Item','SD Item'],'Confirmed Quantity']},
    'SAP MB52':{required:['SD Document','Material'],recommended:['Storage Location','Batch','Unrestricted','Blocked Stock']},
    'UPS':{required:['Tracking Number'],recommended:['Ship Date']},
    'SF Baozun':{required:['Work Order Number','Tracking Number'],recommended:['Ship Group ID','Mother waybill','Qty']},
    'AWB Map':{required:['OrderNumber','ConsolidationShipmentId'],recommended:[]}
  };
  function validateFileRows(parsed){return parsed.map(sheet=>{const schema=schemas[sheet.kind];if(!schema)return {file:sheet.file,sheet:sheet.sheet,kind:sheet.kind,valid:true,missing:[],recommendedMissing:[]};const headers=Object.keys(sheet.rows[0]||{}).map(norm);const missing=schema.required.filter(x=>!headers.includes(norm(x)));const recommendedMissing=schema.recommended.filter(x=>!(Array.isArray(x)?x.some(a=>headers.includes(norm(a))):headers.includes(norm(x)))).map(x=>Array.isArray(x)?x[0]:x);return {file:sheet.file,sheet:sheet.sheet,kind:sheet.kind,valid:missing.length===0,missing,recommendedMissing};});}
  function renderSources() {
    const requiredKinds=['Nike FC','SAP ZRSD1019','SAP MB52','UPS','SF Baozun'];
    for(const kind of [...requiredKinds,'AWB Map']){
      const files=state.files.filter(f=>f.expectedKind===kind||f.types?.includes(kind));
      const el=document.querySelector(`[data-source-status="${kind}"]`);if(!el)continue;
      const usable=files.filter(f=>!f.error&&f.validation?.every(x=>x.valid));
      const multiple=files.length>1;
      el.textContent=!files.length?(kind==='AWB Map'?t('optionalMissing'):t('missingSource')):multiple?t('multipleBatches'):usable.length===files.length?t('ready'):t('warning');
      el.className=`source-readiness ${!files.length?'missing':multiple||usable.length!==files.length?'warning':'ready'}`;
    }
    const root=$('sourceList');
    if(!state.files.length){root.className='source-list empty';root.textContent=t('noFiles');return;}
    root.className='source-list';
    root.innerHTML=state.files.map((f,i)=>{
      const previews=state.sheets.filter(x=>x.file===f.name),v=f.validation||[];
      const validationText=f.error?`${t('importFailed')} · ${uiText(f.error)}`:v.some(x=>!x.valid)?`${t('missingRequired')}：${v.filter(x=>!x.valid).map(x=>`${x.sheet}: ${x.missing.join(', ')}`).join('；')}`:v.some(x=>x.recommendedMissing.length)?t('recommendedMissing',{columns:v.flatMap(x=>x.recommendedMissing).join(', ')}):t('importSuccess');
      const records=previews.reduce((n,x)=>n+x.rows.length,0),cols=previews[0]?Object.keys(previews[0].rows[0]||{}).length:0;
      const dates=previews.flatMap(x=>x.rows.map(row=>cleanDate(value(row,'Manifest Scan Date','Ship Date','Estimated Ship Date'))).filter(Boolean)).sort();
      const dateRange=dates.length?`${dates[0]} ${t('sourceDateRange')} ${dates.at(-1)}`:t('notDetected');
      const preview=previews[0]?.rows.slice(0,10)||[],heads=[...new Set(preview.flatMap(Object.keys))].slice(0,8);
      return `<div class="source-item"><span class="source-name">${escapeHtml(f.name)}</span><span class="source-types"><span class="chip ${f.error?'error':v.some(x=>!x.valid)?'warning':'ok'}">${escapeHtml(f.error?t('readFailed'):v.some(x=>!x.valid)?t('columnWarning'):t('parsed'))}</span><span class="chip">${escapeHtml((f.expectedKind||f.types.join(', '))+' · '+validationText)}</span></span><span class="source-count">${escapeHtml(t('rowsColumns',{rows:records,columns:cols,range:dateRange}))}</span><button type="button" class="source-remove" data-revalidate="${i}">${escapeHtml(t('revalidate'))}</button><button type="button" class="source-remove" data-remove="${i}">${escapeHtml(t('remove'))}</button>${previews.length?`<details class="preview"><summary>${escapeHtml(t('preview'))} · ${previews.map(x=>escapeHtml(x.sheet)).join(', ')}</summary><div class="table-wrap"><table><thead><tr>${heads.map(h=>`<th>${escapeHtml(displayText(h))}</th>`).join('')}</tr></thead><tbody>${preview.map(row=>`<tr>${heads.map(h=>`<td>${escapeHtml(displayText(row[h]??''))}</td>`).join('')}</tr>`).join('')}</tbody></table></div></details>`:''}</div>`;
    }).join('');
  }
  function renderSourceChoices(){
    const select=$('fcSource');
    const previous=select.value;
    const choices=state.sheets.filter(s=>s.kind==='Nike FC').map(s=>{
      const rows=s.rows.length.toLocaleString();
      const dates=s.rows.map(r=>cleanDate(value(r,'Manifest Scan Date'))).filter(Boolean).sort();
      const range=dates.length?`${dates[0]} ${t('sourceDateRange')} ${dates.at(-1)}`:t('notDetected');
      return `<option value="${escapeHtml(`${s.file}::${s.sheet}`)}">${escapeHtml(`${s.file} · ${s.sheet} · ${t('rowsCount',{count:rows})} · ${range}`)}</option>`;
    });
    select.innerHTML=`<option value="" data-i18n="importFirst">${escapeHtml(t('importFirst'))}</option>${choices.join('')}`;
    if(previous&&choices.length&&[...select.options].some(o=>o.value===previous))select.value=previous;else if(choices.length===1)select.value=select.options[1].value;
  }
  function validSourceSheets(kind){return state.sheets.filter(s=>s.kind===kind&&state.files.find(f=>f.name===s.file)?.validation?.find(v=>v.sheet===s.sheet)?.valid!==false);}
  function recommendedSourceBook(kind,sheets){
    const books=[...new Set(sheets.map(s=>s.file))];if(books.length<2)return books[0]||'';
    const filePatterns={'SAP ZRSD1019':/(?:^|[^A-Z0-9])(?:SAP[ _.-]*)?ZRSD1019(?:$|[^A-Z0-9])/i,'SAP MB52':/^MB52(?:$|[ _.\-])/i,'UPS':/(?:^|[^A-Z0-9])UPS(?:$|[^A-Z0-9])|DETAILED[ _-]*REPORT/i,'SF Baozun':/SF|BAOZUN|SHIPMENT/i,'AWB Map':/AWB|CONSOLIDATION/i};
    const sheetPatterns={'SAP ZRSD1019':/ZRSD1019/i,'SAP MB52':/MB52/i,'UPS':/UPS|DETAILED/i,'SF Baozun':/SF|BAOZUN|MANIFEST|SHIPMENT/i,'AWB Map':/AWB|MAPPING|CONSOLIDATION/i};
    const byName=books.filter(file=>filePatterns[kind]?.test(file.split(/[\/]/).pop()||file));if(byName.length===1)return byName[0];
    const bySheet=books.filter(file=>sheets.some(s=>s.file===file&&sheetPatterns[kind]?.test(s.sheet)));return bySheet.length===1?bySheet[0]:'';
  }
  function renderSourceOverrides(){
    const root=$('sourceOverrides');if(!root)return;
    const kinds=['SAP ZRSD1019','SAP MB52','UPS','SF Baozun','AWB Map'];
    const entries=kinds.map(kind=>({kind,sheets:validSourceSheets(kind)})).filter(x=>new Set(x.sheets.map(s=>s.file)).size>1);
    root.hidden=!entries.length;if(!entries.length){root.innerHTML='';return;}
    const controls=entries.map(({kind,sheets})=>{
      const books=[...new Set(sheets.map(s=>s.file))];let chosen=state.sourceOverrides[kind];
      if(chosen&&!books.includes(chosen)){delete state.sourceOverrides[kind];chosen='';}
      const manuallyChosen=!!chosen;
      chosen=chosen||recommendedSourceBook(kind,sheets);
      const optionText=file=>{const tabs=sheets.filter(s=>s.file===file).map(s=>s.sheet).join(', ');return `${file} · ${tabs}`;};
      return `<label class="source-override-label"><span>${escapeHtml(t('sourceSelectionLabel',{kind}))}</span><select data-source-override="${escapeHtml(kind)}"><option value="">${escapeHtml(t('chooseSourceBook',{kind}))}</option>${books.map(file=>`<option value="${escapeHtml(file)}" ${file===chosen?'selected':''}>${escapeHtml(optionText(file))}</option>`).join('')}</select><small>${escapeHtml(chosen?t(manuallyChosen?'sourceManualSelected':'sourceAutoSelected'):t('chooseSourceBook',{kind}))}</small></label>`;
    }).join('');
    root.innerHTML=`<div class="source-override-card"><strong>${escapeHtml(t('sourceOverridesTitle'))}</strong><p>${escapeHtml(t('sourceOverridesHint'))}</p><div class="source-override-grid">${controls}</div></div>`;
  }
  $('sourceList').addEventListener('click',e=>{const re=e.target.closest('[data-revalidate]');if(re&&!state.busy){const f=state.files[Number(re.dataset.revalidate)],rows=state.sheets.filter(s=>s.file===f.name);if(!rows.length&&f.sourceFile){ingest([f.sourceFile],f.expectedKind||'');return;}f.validation=validateFileRows(rows);f.error=rows.length?'':'沒有可重新驗證的資料表';renderSources();renderSourceChoices();renderSourceOverrides();return;}const b=e.target.closest('[data-remove]');if(!b||state.busy)return;const f=state.files[Number(b.dataset.remove)];state.files.splice(Number(b.dataset.remove),1);state.sheets=state.sheets.filter(s=>s.file!==f.name);invalidate();renderSources();renderSourceChoices();renderSourceOverrides();$('runButton').disabled=!state.sheets.length;});
  function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
  async function calculate() {
    const target=dateForInput();
    if(!/^\d{4}-\d{2}-\d{2}$/.test(target))throw Error(t('validDate'));
    const all=k=>state.sheets.filter(s=>s.kind===k&&state.files.find(f=>f.name===s.file)?.validation?.find(v=>v.sheet===s.sheet)?.valid!==false).flatMap(s=>s.rows);
    const fcCandidates=state.sheets.filter(s=>s.kind==='Nike FC');
    const sourceDecisions={};const source=k=>{const ss=validSourceSheets(k),books=[...new Set(ss.map(s=>s.file))];let chosen=books.length===1?books[0]:state.sourceOverrides[k]||recommendedSourceBook(k,ss);if(books.length>1&&!chosen)throw Error(t('chooseSourceBook',{kind:k}));if(chosen)sourceDecisions[k]={selected:chosen,ignored:books.filter(file=>file!==chosen)};return NBYEngine.unique(ss.filter(s=>!chosen||s.file===chosen).flatMap(s=>s.rows));};
    const sap=source('SAP ZRSD1019'), mb52Raw=source('SAP MB52'), ups=source('UPS'), sf=source('SF Baozun').filter(r=>!(norm(value(r,'Factory'))==='TOTAL'&&!value(r,'Work Order Number')&&!value(r,'Tracking Number')));
    const mb52=NBYEngine.unique(mb52Raw);
    const rawFcSheets=validSourceSheets('Nike FC').filter(s=>!s.processed);
    if(!rawFcSheets.length)throw Error(t('fcMissing'));
    const previousDay=new Date(target+'T00:00:00Z');previousDay.setUTCDate(previousDay.getUTCDate()-1);
    const reportDates={orders:previousDay.toISOString().slice(0,10),data:target,statusStart:target.slice(0,8)+'01',statusEnd:target};
    const dateOf=row=>cleanDate(value(row,'Manifest Scan Date'));
    const rawRows=NBYEngine.unique(rawFcSheets.flatMap(s=>s.rows));
    const selectedOrders=rawRows.filter(row=>dateOf(row)===reportDates.orders);
    const selectedData=rawRows.filter(row=>dateOf(row)===reportDates.data);
    const selectedStatus=rawRows.filter(row=>dateOf(row)&&dateOf(row)<=target);
    const fcRows=NBYEngine.unique([...selectedOrders,...selectedData,...selectedStatus]);
    const selected={file:[...new Set(rawFcSheets.map(s=>s.file))].join(', '),sheet:'Multiple raw FC worksheets',rows:fcRows,processed:false};
    const selectedFc={s:selected,rows:fcRows,ordersRows:selectedOrders,dataRows:selectedData,statusRows:selectedStatus};
    const manualSheets=state.sheets.filter(s=>s.kind==='人工成品');
    const {orders,upsRows,sfRows,anomalies}=NBYEngine.reconcile({fc:selectedData,sap,stock:mb52,ups,sf,awb:source('AWB Map'),date:target,cleanDate});
    const add=(severity,category,key,detail)=>anomalies.push({嚴重度:severity,異常類型:category,工單或Tracking:key,說明:detail});
    if(!selectedOrders.length)add('中','Orders Update — Missing Required Source',reportDates.orders,t('missingPreviousFc',{date:reportDates.orders}));
    selectedFc.linked=orders.filter(o=>o.sapMatch).length;selectedFc.rate=orders.length?selectedFc.linked/orders.length:0;

    for(const f of state.files)if(f.error||!f.sheetCount)add('高','來源檔讀取失敗或未辨識',f.name,f.error||'沒有可辨識資料表');else if(f.validation?.some(v=>!v.valid))add('高','來源缺少必要欄位',f.name,f.validation.filter(v=>!v.valid).map(v=>`${v.sheet}: ${v.missing.join(', ')}`).join('；'));
    for(const kind of ['SAP ZRSD1019','SAP MB52','UPS','SF Baozun'])if(!all(kind).length)add('中','來源未提供或未通過欄位驗證',kind,'相關報表狀態會標示缺少必要來源；不以空白欄位冒充匹配');
    if(!all('AWB Map').length)add('低','AWB Mapping 未提供','AWB Map','AWB 保留缺值，不自行產生。');
    if(selected.processed)add('高','使用人工加工資料作為 FC',selected.file,'可測試關聯，但不可視為獨立 FC 原始資料 Golden Sample 驗收通過');
    const selectedDates=[...new Set(fcRows.map(dateOf).filter(Boolean))];
    if(!/^\d{4}-\d{2}-\d{2}$/.test(target))throw Error('請選擇有效日期');
    const manual=validateManual(manualSheets,orders,upsRows,sfRows,mb52);

    state.result={target,reportDates,sourceDecisions,orders,upsRows,sfRows,anomalies,manual,counts:{fcRaw:selectedFc?.s.rows.length||0,fcDate:fcRows.length,orders:orders.length,ordersRows:selectedOrders.length,dataRows:selectedData.length,statusRows:selectedStatus.length,sap:sap.length,mb52:mb52.length,ups:ups.length,sf:sf.length,manualFiles:new Set(manualSheets.map(x=>x.file)).size},selectedFc:selectedFc?{...selectedFc.s,linked:selectedFc.linked,rate:selectedFc.rate,rows:selectedFc.rows,ordersRows:selectedFc.ordersRows,dataRows:selectedFc.dataRows,statusRows:selectedFc.statusRows}:null};
    await buildOutputs(); renderResults();
  }
  function validateManual(sheets,orders,upsRows,sfRows,mb52){
    const summaries=[]; const lines=[]; const byWO=new Map(orders.map(o=>[o.wo,o]));
    const byTrack=new Map(upsRows.map(x=>[norm(x.Tracking_Number),x])); const sfByWO=new Map(sfRows.map(x=>[norm(x.Work_Order_Number),x]));
    const mbKey=(doc,item,material)=>[norm(doc),norm(item),norm(material)].join('|');
    const mbSet=new Set(mb52.map(r=>mbKey(value(r,'SD Document'),value(r,'Item (SD)','SD Item'),value(r,'Material'))));
    const summary=(s,type,count,ok,details='')=>summaries.push({檔案:s.file,工作表:s.sheet,比對項目:type,人工筆數:count,成功比對:ok,結果:count===ok?'通過':'需檢查',差異摘要:details});
    for(const s of sheets){
      const rows=s.rows;
      const baselineRows=rows.filter(r=>value(r,'Work Order Nbr')&&value(r,'Manifest Scan Date'));
      if(baselineRows.length){
        const hasDataColumns=Object.keys(baselineRows[0]).some(k=>norm(k)==='ZRSD1019');
        if(!hasDataColumns || /\(2\)/.test(s.sheet)){
          summary(s,'人工報表驗收未執行',baselineRows.length,0,'缺少獨立來源或此工作表為中間加工階段');
          continue;
        }
        let matched=0,checks=0,matches=0;
        for(const row of baselineRows){
          const wo=norm(value(row,'Work Order Nbr')), order=byWO.get(wo), differences=[];
          if(!order) differences.push('人工成品有／轉檔結果無');
          else {
            const expected={...order.source,ZRSD1019:order.sapDoc,MB52:order.mbDoc,AWB:order.awb};
            for(const [column,manual] of Object.entries(row)){
              const key=Object.keys(expected).find(k=>norm(k)===norm(column));
              if(!key||/^欄位\d/.test(column))continue;
              checks++;
              const normalise=v=>/date/i.test(column)?cleanDate(v):norm(v);
              const left=normalise(manual),right=normalise(expected[key]);
              if(left===right)matches++;else differences.push(`${column}:人工=${manual}／系統=${expected[key]??''}`);
            }
          }
          if(!differences.length)matched++;
          lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'人工報表欄位',工單:wo,結果:differences.length?'有差異':'通過',差異:differences.join('；')});
        }
        summary(s,'人工報表逐筆／欄位',baselineRows.length,matched,`欄位相符 ${matches}/${checks}`);
        continue;
      }
      const orderRows=rows.filter(r=>value(r,'Work Order Nbr')&&(value(r,'SAP 訂單匹配')||value(r,'SAP SD Document ZRSD')||value(r,'核對狀態')));
      if(orderRows.length){let matched=0,fieldChecks=0,fieldMatches=0;
        for(const r of orderRows){const wo=norm(value(r,'Work Order Nbr'));const o=byWO.get(wo);let diffs=[];let rowChecks=0;
          if(!o)diffs.push('人工成品有／轉檔結果無');else{
            const checks=[['FC Scan Date',o.fcDate,v=>cleanDate(v)],['Shipper',norm(o.shipper),v=>norm(v)],['Tracking Number',o.track,v=>norm(v)],['SAP SD Document ZRSD',norm(o.sapDoc),v=>norm(v)],['SAP SD Document MB52',norm(o.mbDoc),v=>norm(v)],['SAP 訂單匹配',o.sapMatch?'YES':'NO',v=>norm(v)],['MB52 庫存紀錄匹配',o.stockMatch?'YES':'NO',v=>norm(v)],['UPS Tracking 匹配',NBYEngine.carrier(o.shipper)==='UPS'?(o.upsMatch?'YES':'NO'):'N/A',v=>norm(v)],['AWB / Consolidation ID',norm(o.awb),v=>norm(v)],['AWB 狀態',o.awb?'有':'未填',v=>norm(v)],['核對狀態','工單及來源已匹配（非數量核對）',v=>String(v??'').trim()]];
            for(const [col,expected,coerce] of checks){const got=value(r,col);if(got!==''&&got!=null){fieldChecks++;rowChecks++;if(coerce(got)===coerce(expected))fieldMatches++;else diffs.push(`${col}:人工=${got}／系統=${expected}`);}}
          }
          if(!diffs.length&&rowChecks===0)diffs.push('沒有可比對欄位');
          if(!diffs.length)matched++;
          lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'工單欄位',工單:wo,結果:diffs.length?'有差異':'通過',差異:diffs.join('；'),欄位已比對:rowChecks});
        }
        summary(s,'工單逐筆／欄位',orderRows.length,matched,`欄位相符 ${fieldMatches}/${fieldChecks}`);
      }
      const trackRows=rows.filter(r=>value(r,'Tracking Number')&&(value(r,'UPS 原始出現次數')!==''||value(r,'FC Data Shipping 出現次數')!==''||value(r,'紀錄次數')!==''||value(r,'重複筆數')!==''));
      if(trackRows.length){let matched=0;
        for(const r of trackRows){const t=norm(value(r,'Tracking Number'));const actual=byTrack.get(t);let diffs=[];if(!actual)diffs.push('人工 Tracking 不在 UPS 原始資料');else{
          for(const [col,expected] of [['UPS 原始出現次數',actual.UPS_紀錄數],['紀錄次數',actual.UPS_紀錄數],['重複筆數',actual.UPS_紀錄數],['FC Data Shipping 出現次數',actual.FC_工單數]]){const got=value(r,col);if(got!==''&&got!=null&&norm(got)!==norm(expected))diffs.push(`${col}:人工=${got}／系統=${expected}`);}
          const service=value(r,'UPS Service','Service');if(service&&norm(service)!==norm(actual.UPS_Service))diffs.push(`Service:人工=${service}／系統=${actual.UPS_Service}`);
          const charge=value(r,'Charge');if(charge!==''&&norm(charge)!==norm(actual.UPS_Charge))diffs.push(`Charge:人工=${charge}／系統=${actual.UPS_Charge}`);
          const ship=value(r,'Ship Date');if(ship&&cleanDate(ship)!==actual.Ship_Date)diffs.push(`Ship Date:人工=${ship}／系統=${actual.Ship_Date}`);
        }
          if(!diffs.length)matched++;lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'UPS Tracking',Tracking_Number:t,結果:diffs.length?'有差異':'通過',差異:diffs.join('；')});
        }
        summary(s,'UPS Tracking 逐筆',trackRows.length,matched);
      }
      const sfManual=rows.filter(r=>value(r,'Work Order Number')&&value(r,'Mother waybill')&&value(r,'Tracking Number'));
      if(sfManual.length){let matched=0;
        for(const r of sfManual){const wo=norm(value(r,'Work Order Number'));const actual=sfByWO.get(wo);let diffs=[];if(!actual)diffs.push('人工 SF 工單不在 manifest');else{
          for(const [col,got,want] of [['Tracking Number',value(r,'Tracking Number'),actual.Tracking_Number],['Ship Group ID',value(r,'Ship Group ID'),actual.Ship_Group_ID],['Mother waybill',value(r,'Mother waybill'),actual.Mother_Waybill],['Qty',value(r,'Qty'),actual.Qty]])if(got!==''&&norm(got)!==norm(want))diffs.push(`${col}:人工=${got}／系統=${want}`);
          const d=value(r,'Estimated Ship Date','Est Ship Date');if(d&&cleanDate(d)!==actual.Estimated_Ship_Date)diffs.push(`Ship Date:人工=${d}／系統=${actual.Estimated_Ship_Date}`);
        }
          if(!diffs.length)matched++;lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'SF Baozun manifest',工單:wo,結果:diffs.length?'有差異':'通過',差異:diffs.join('；')});
        }
        summary(s,'SF Baozun 逐筆',sfManual.length,matched);
      }
      const stockRows=rows.filter(r=>value(r,'SD Item','Item (SD)')&&value(r,'Material'));
      if(stockRows.length){let matched=0;
        for(const r of stockRows){const exists=mbSet.has(mbKey(value(r,'SD Document'),value(r,'SD Item','Item (SD)'),value(r,'Material')));if(exists)matched++;lines.push({人工檔案:s.file,工作表:s.sheet,比對類型:'MB52 庫存列',SD_Document:value(r,'SD Document'),SD_Item:value(r,'SD Item','Item (SD)'),Material:value(r,'Material'),結果:exists?'通過':'原始來源無此列',差異:exists?'':'文件／項次／物料組合未匹配'});}
        summary(s,'MB52 庫存列逐筆',stockRows.length,matched);
      }
    }
    if(!summaries.length)summaries.push({檔案:'',工作表:'',比對項目:'人工成品',人工筆數:0,成功比對:0,結果:sheets.length?'已匯入人工成品，但沒有可比對的工作表':'尚未匯入人工成品'});
    return {summaries,lines};
  }
  const metricKeys=['metricOrders','metricData','metricStatus','metricSap','metricMb','metricExceptions'];
  function renderResults(){
    const r=state.result;$('results').classList.remove('hidden');
    const missing=r.anomalies.filter(x=>x.異常類型==='Orders Update — Missing Required Source').length;
    const shipment=r.anomalies.length-missing;
    const validationDiff=r.manual.lines.filter(x=>x.結果!=='通過').length;
    const labels={zh:['出貨異常','必要來源缺漏','人工報表驗收差異','系統操作資訊'],en:['Shipment exceptions','Missing required sources','Manual report differences','Processing information'],vi:['Ngoại lệ giao hàng','Thiếu nguồn bắt buộc','Chênh lệch báo cáo thủ công','Thông tin xử lý']}[language];
    const m=[[labels[0],shipment,''],[labels[1],missing,''],[labels[2],validationDiff,''],[labels[3],Object.keys(r.sourceDecisions).length,''],[t('metricOrders'),r.counts.ordersRows,r.reportDates.orders],[t('metricData'),r.counts.dataRows,r.reportDates.data],[t('metricStatus'),r.counts.statusRows,t('statusUnit')],[t('metricSap'),r.counts.sap,t('statusSap')],[t('metricMb'),r.counts.mb52,t('statusMb')]];
    $('metrics').innerHTML=m.map(x=>`<div class="metric"><span>${escapeHtml(x[0])}</span><strong>${escapeHtml(x[1])}</strong><small>${escapeHtml(x[2])}</small></div>`).join('')+`<div class="source-note"><strong>${escapeHtml(t('sourceNote'))}</strong><span>${escapeHtml(r.selectedFc?.file||t('genericSource'))} · ${escapeHtml(r.selectedFc?.sheet||'')} · ${escapeHtml(t('sourceMatch',{count:r.counts.orders,rate:r.counts.orders?Math.round(r.selectedFc.linked/Math.max(1,r.counts.orders)*100):0}))}</span><small>${escapeHtml(t('sourceNoteText'))}</small></div>`;
    renderTable();
  }
  const tabData=()=>{const r=state.result;switch(r.activeTab||state.activeTab){case'ups':return r.upsRows;case'sf':return r.sfRows;case'anomalies':return r.anomalies;case'manual':return r.manual.lines.length?r.manual.lines:r.manual.summaries.map(({details,...x})=>x);default:return r.orders.map(o=>({'Work Order Nbr':o.wo,'FC Scan Date':o.fcDate,Shipper:o.shipper,'Tracking Number':o.track,'SAP SD Document ZRSD':o.sapDoc,'SAP SD Document MB52':o.mbDoc,'SAP 訂單匹配':o.sapMatch?'YES':'NO','MB52 庫存紀錄匹配':o.stockMatch?'YES':'NO','UPS Tracking 匹配':o.upsMatch?'YES':'','AWB / Consolidation ID':o.awb,'AWB 狀態':o.awb?'有':'未填'}));}};
  function renderTable(){if(!state.result)return;const data=tabData();const q=norm($('tableSearch').value);const filtered=data.filter(row=>!q||Object.values(row).some(v=>norm(v).includes(q)));const cols=[...new Set(filtered.flatMap(Object.keys))].filter(c=>c!=='details');$('tableHead').innerHTML=`<tr>${cols.map(c=>`<th>${escapeHtml(displayText(c))}</th>`).join('')}</tr>`;$('tableBody').innerHTML=filtered.slice(0,500).map(row=>`<tr>${cols.map(c=>`<td>${formatCell(row[c])}</td>`).join('')}</tr>`).join('');$('tableCount').textContent=filtered.length>500?t('tableRowsLimited',{count:filtered.toLocaleString()}):t('tableRows',{count:filtered.toLocaleString()});}
  function displayText(v){const s=String(v??'');return phraseMaps[language]?.[s]??s;}
  function formatCell(v){if(typeof v==='boolean')return `<span class="state ${v?'pass':'fail'}">${escapeHtml(v?t('yes'):t('no'))}</span>`;const s=String(v??''),txt=displayText(s);if(['通過','已與 FC Tracking 對上','FC Tracking 一致','FC Tracking 一致；數量未認證'].includes(s))return `<span class="state pass">${escapeHtml(txt)}</span>`;if(/未匹配|不一致|重複|缺少|待檢查/.test(s))return `<span class="state warn">${escapeHtml(txt)}</span>`;return escapeHtml(txt);}
  const outputNames=['Orders Update.xlsx','Data Shipping.xlsx','Shipping Status.xlsx','UPS Reconciliation.xlsx','SF Baozun Reconciliation.xlsx','Exception Report.xlsx','Golden Sample Difference Report.xlsx'];
  async function buildOutputs(){
    const r=state.result;if(!r)return {};
    const cleanRows=rows=>(rows||[]).map(row=>Object.fromEntries(Object.entries(row).map(([k,v])=>[k,v instanceof Date?cleanDate(v):v])));
    const sumNumeric=(rows,column)=>{const values=rows.map(row=>NBYEngine.qty(value(row,column)));return values.length&&values.every(v=>v!==null)?values.reduce((a,b)=>a+b,0):'';};
    const outputHeaderMap={'SAP_Confirmed_Qty_Reference':'SAP Confirmed Quantity (Reference)','Qty_vs_SAP_Reference':'SF Qty vs SAP Reference','MB52_Unrestricted_Reference':'MB52 Unrestricted (Reference)','MB52_Blocked_Stock_Reference':'MB52 Blocked Stock (Reference)','SF_Qty_minus_MB52_Unrestricted_Reference':'SF Qty minus MB52 Unrestricted (Reference)','MB52_Interpretation':'MB52 Interpretation','數量核對':'Quantity Validation Status','工單或Tracking':'Work Order or Tracking Number','嚴重度':'Severity','異常類型':'Exception Type','說明':'Details','檔案':'File','人工檔案':'Manual File','工作表':'Worksheet','比對類型':'Check Type','工單':'Work Order','結果':'Result','差異':'Difference','欄位已比對':'Fields Compared','人工筆數':'Manual Rows','成功比對':'Matched Rows','比對項目':'Check Item','差異摘要':'Difference Summary','UPS_紀錄數':'UPS Record Count','FC_工單數':'FC Work Order Count','FC_工單':'FC Work Orders','MB52_明細列數':'MB52 Detail Rows','SAP_SD_Document':'SAP SD Document','SAP_Confirmed_Quantity':'SAP Confirmed Quantity','MB52_Unrestricted':'MB52 Unrestricted Stock','MB52_Blocked_Stock':'MB52 Blocked Stock','UPS_Service':'UPS Service','Ship_Date':'Ship Date','UPS_Charge':'UPS Charge','Work_Order_Number':'Work Order Number','Ship_Group_ID':'Ship Group ID','Mother_Waybill':'Mother Waybill','Estimated_Ship_Date':'Estimated Ship Date','SAP_Confirmed_Qty':'SAP Confirmed Quantity','UPS_原始出現次數':'UPS Source Occurrences','紀錄次數':'Record Count','重複筆數':'Duplicate Count','SAP_訂單匹配':'SAP Order Matched','核對狀態':'Reconciliation Status','MB52_庫存紀錄匹配':'MB52 Stock Row Matched','UPS_Tracking_匹配':'UPS Tracking Matched','AWB_狀態':'AWB Status','Qty與SAP一致':'SF Qty vs SAP (Reference Only)','SAP_工單存在':'SAP Work Order Found','MB52_紀錄存在':'MB52 Record Found','FC_工單存在':'FC Work Order Found','FC_Tracking一致':'FC Tracking Matches'};
    const outputTextMap={'高':'High','中':'Medium','低':'Low','通過':'PASS','有差異':'DIFFERENCE','人工成品逐筆比對差異':'Manual output row-by-row difference','FC 工單缺少 SAP ZRSD1019':'FC work order missing from SAP ZRSD1019','SAP 訂單缺少 MB52 庫存紀錄':'SAP order missing from MB52 stock snapshot','FC Tracking 未出現在 UPS':'FC Tracking Number missing from UPS source','FC 缺少 Tracking Number':'FC work order has no Tracking Number','未找到 AWB／Consolidation ID':'AWB / Consolidation ID not found','UPS Tracking 重複':'Duplicate UPS Tracking Number','UPS Tracking 未關聯 FC 工單':'UPS Tracking Number not linked to FC work order','SF manifest 工單未匹配 FC':'SF manifest work order not matched to FC','SF Qty 與 SAP 數量不一致':'SF quantity differs from SAP confirmed quantity (reference only)','SF Qty 與 SAP Confirmed Qty 參考差異（待判讀）':'SF quantity differs from SAP confirmed quantity (reference only; review pending)','MB52 庫存快照差額（待業務判讀）':'MB52 stock snapshot delta (business interpretation pending)','SF／FC Tracking 不一致':'SF and FC Tracking Numbers do not match','FC 工單重複':'Duplicate FC work order','FC 缺少 Work Order Nbr':'FC rows missing Work Order Nbr','依 Customer Reference 找不到對應工單':'No work order found by Customer Reference','無法執行 UPS 逐筆核對':'UPS row-level reconciliation could not run','來源未提供對應 AWB 或 Consolidation ID':'No AWB or Consolidation ID is available in the source','未建立':'Not created','未提供':'Not provided','空白':'Blank','有':'Present','未填':'Blank','工單及來源已匹配（非數量核對）':'Work order and source rows matched (not a quantity check)','原始來源無此列':'Source row not found','人工成品有／轉檔結果無':'Present in manual output but missing from generated output','人工 Tracking 不在 UPS 原始資料':'Tracking Number not found in UPS source','一致（Tracking 存在）':'MATCH (Tracking Number found)','UPS 重複紀錄':'Duplicate UPS record; review required','已與 FC Tracking 對上':'Matched to FC Tracking Number','UPS 有紀錄但 FC 無對應':'UPS record has no matching FC row','SF 工單未出現在 FC':'SF Work Order not found in FC','SF Qty 與 SAP 數量不一致':'SF quantity differs from SAP confirmed quantity (reference only)','SF Qty 與 SAP Confirmed Qty 參考差異（待判讀）':'SF quantity differs from SAP confirmed quantity (reference only; review pending)','MB52 庫存快照差額（待業務判讀）':'MB52 stock snapshot delta (business interpretation pending)','Tracking 不一致':'Tracking Number mismatch','FC Tracking 一致；數量未認證':'FC Tracking matches; quantity not certified','FC Tracking 一致；SAP Confirmed Qty 僅供參考，MB52 數量待判讀':'FC Tracking matches; SAP confirmed quantity is reference only; MB52 needs business interpretation','FC Tracking 一致':'FC Tracking matches (tracking only)','工單及來源已匹配（非數量核對）':'Work order and source records matched (not a quantity check)','尚未匯入人工成品':'No manual output file imported','未核對 FC / SAP':'FC / SAP not checked','NBY_1120_完整源頭逐筆比對.xlsx':'NBY_1120_Full_Source_Row_Comparison.xlsx','03_工單逐筆對照':'03_Work_Order_Row_Comparison','04_UPS_異常核對':'04_UPS_Exceptions','UPS 出現 2 筆；需確認是否為分件或重複下載':'UPS contains multiple records; confirm split shipment versus duplicate download','來源未提供對應 AWB 或 Consolidation ID':'No AWB or Consolidation ID is available in the source','物流明細存在，但本次日期 FC 工單中無相同 Tracking Number':'UPS record exists but no matching Tracking Number is present in FC for the selected date','工單欄位':'Work Order Fields','人工成品':'Manual Output','工單逐筆／欄位':'Work Order Row and Field Check','UPS Tracking 逐筆':'UPS Tracking Row Check','欄位相符':'Fields matched','NBY_1120_完整源頭逐筆比對.xlsx':'NBY_1120_Full_Source_Row_Comparison.xlsx','03_工單逐筆對照':'03_Work_Order_Row_Comparison','04_UPS_異常核對':'04_UPS_Exceptions'};
    const englishText=v=>{let s=String(v??'');if(outputTextMap[s])return outputTextMap[s];s=s.replace(/^SAP SD Document (.+?) 在 MB52 無紀錄$/,'SAP SD Document $1 has no matching row in MB52').replace(/^FC 工單 (.+?) 的 Tracking Number 未在 UPS 清單找到$/,'FC work order $1: Tracking Number not found in the UPS source').replace(/^UPS 出現 (\d+) 筆；需確認是否為分件或重複下載$/,'UPS contains $1 records; confirm split shipment versus duplicate download').replace(/^欄位相符 (.+)$/,'Fields matched $1').replace(/^原始資料 (\d+) 筆工單號重複$/,'$1 repeated Work Order Nbr records in source').replace(/([^:／；]+):人工=/g,(m,k)=>`${outputHeaderMap[k]||outputTextMap[k]||k}:Manual=`).replace(/／系統=/g,' / Generated=').replace(/人工=/g,'Manual=');return outputTextMap[s]||s;};
    const englishRows=rows=>cleanRows(rows).map(row=>Object.fromEntries(Object.entries(row).map(([k,v])=>[outputHeaderMap[k]||k,englishText(v)])));
    const headersOf=rows=>[...new Set((rows||[]).flatMap(Object.keys))];
    const matchByWo=new Map(r.orders.map(o=>[norm(o.wo),o]));
    const orderSource=r.selectedFc;
    const orderRows=cleanRows(orderSource?.ordersRows||[]);
    const dataRows=cleanRows((orderSource?.dataRows||[]).map(src=>{
      const wo=norm(value(src,'Work Order Nbr')),o=matchByWo.get(wo)||{};
      const row={...src};
      row.ZRSD1019=o.sapDoc||'';
      row.MB52=o.mbDoc||'';
      row.AWB=o.awb||'';
      return row;
    }));
    const statusRows=cleanRows(orderSource?.statusRows||[]);
    const historicalDates=[],historicalProducts=new Map();
    const historySummary=state.sheets.filter(s=>s.kind==='Shipping Status Summary').flatMap(s=>s.rows);
    const historyDetail=state.sheets.filter(s=>s.kind==='Shipping Status Detail').flatMap(s=>s.rows);
    for(const row of historyDetail){const product=String(value(row,'Product Name')||'').trim();if(!product)continue;if(!historicalProducts.has(product))historicalProducts.set(product,{});for(const [header,count] of Object.entries(row)){if(norm(header)==='PRODUCT NAME'||count===''||count==null)continue;const date=cleanDate(header);if(!date||date===header&&/[^0-9./-]/.test(header))continue;if(!historicalDates.includes(date))historicalDates.push(date);historicalProducts.get(product)[date]=count;}}
    const productDateOrders=new Map(),totalDateOrders=new Map();
    for(const row of statusRows){const date=cleanDate(value(row,'Manifest Scan Date')),product=String(value(row,'Product Name')||'Not provided'),wo=norm(value(row,'Work Order Nbr'));if(!date||!wo)continue;const pk=`${date}|${product}`;if(!productDateOrders.has(pk))productDateOrders.set(pk,{date,product,workOrders:new Set()});productDateOrders.get(pk).workOrders.add(wo);if(!totalDateOrders.has(date))totalDateOrders.set(date,new Set());totalDateOrders.get(date).add(wo);}
    const computedDates=[...totalDateOrders.keys()].sort();
    const addDay=date=>{const d=new Date(date+'T00:00:00Z');d.setUTCDate(d.getUTCDate()+1);return d.toISOString().slice(0,10);};
    const pairCount=(workOrders,date)=>date===r.reportDates.data&&[...workOrders].every(wo=>matchByWo.get(wo)?.pairQty!=null)?[...workOrders].reduce((total,wo)=>total+matchByWo.get(wo).pairQty,0):null;
    for(const x of productDateOrders.values()){if(!historicalProducts.has(x.product))historicalProducts.set(x.product,{});historicalProducts.get(x.product)[x.date]=x.workOrders.size;if(!historicalDates.includes(x.date))historicalDates.push(x.date);}
    const statusSummary=historySummary.filter(row=>!computedDates.includes(cleanDate(value(row,'Date')))).map(row=>({'Date':cleanDate(value(row,'Date')),'Actually Shipment':value(row,'Actually Shipment'),'Expect Shipment':value(row,'Expect Shipment')}));
    for(const date of computedDates)statusSummary.push({'Date':date,'Actually Shipment':totalDateOrders.get(date).size,'Expect Shipment':historySummary.find(row=>cleanDate(value(row,'Date'))===date)?.['Expect Shipment']??'','SAP Verified PAA':pairCount(totalDateOrders.get(date),date),'Actually Shipment Rule':'Manifest work orders; SAP PAA verified separately','Expect Shipment Rule':'PENDING — not calculated'});
    for(const date of computedDates){
      const pairs=pairCount(totalDateOrders.get(date),date);if(pairs===null)continue;
      const next=addDay(date);let summary=statusSummary.find(row=>row.Date===next);
      if(!summary){summary={Date:next,'Actually Shipment':'','Expect Shipment':''};statusSummary.push(summary);}
      summary['Expect Shipment']=pairs;
      summary['Expect Shipment Rule']='Previous day Manifest; SAP PAA verified';
    }
    const statusDetail=[...historicalProducts].map(([product,counts])=>Object.assign({'Product Name':product},...historicalDates.sort().map(date=>({[date]:counts[date]??''}))));
    const upsBook=r.sourceDecisions?.UPS?.selected;const upsRaw=cleanRows(state.sheets.filter(s=>s.kind==='UPS'&&(!upsBook||s.file===upsBook)).flatMap(s=>s.rows));
    const upsExport=upsRaw.map((row,i)=>{const track=norm(value(row,'Tracking Number')),ordersForTrack=r.orders.filter(o=>norm(o.shipper)==='UPS'&&norm(o.track)===track);return {...row,'FC Work Order Nbr':ordersForTrack.map(o=>o.wo).join(', '),'FC Match Count':ordersForTrack.length,'Duplicate Tracking in UPS':(r.upsRows.find(u=>norm(u.Tracking_Number)===track)?.UPS_紀錄數||0)>1?'REVIEW':'','Source File':state.sheets.find(s=>s.kind==='UPS')?.file||''};});
    const sfExport=englishRows(r.sfRows);
    const exceptions=englishRows(r.anomalies);
    const goldenDifferences=r.manual.lines.flatMap(line=>{const key=line.工單||line.Tracking_Number||line.SD_Document||'';const bits=String(line.差異||'').split('；').filter(Boolean);if(!bits.length)return [{Report_Name:line.比對類型||'Manual Output',Worksheet:line.工作表||'',Work_Order_or_Record_Key:key,Column_Name:'Checked fields',Manual_Value:'',Generated_Value:'',Difference:'',Source_Reference:`${line.人工檔案||''} / ${line.工作表||''} / ${key}`,Validation_Result:line.結果==='通過'?'PASS – Verified':'WARNING – Review'}];return bits.map(bit=>{const m=bit.match(/^([^:]+):人工=(.*?)／系統=(.*)$/);return {Report_Name:line.比對類型||'Manual Output',Worksheet:line.工作表||'',Work_Order_or_Record_Key:key,Column_Name:m?.[1]||'Record Presence',Manual_Value:m?.[2]??bit,Generated_Value:m?.[3]??'',Difference:bit,Source_Reference:`${line.人工檔案||''} / ${line.工作表||''} / ${key}`,Validation_Result:/缺少|不在|無此列/.test(bit)?'FAIL – Incorrect':'WARNING – Needs Review'};});});
    const testSummary=[
      {Check:'Orders Update source rows',Rule:`Filter by Manifest Scan Date ${r.reportDates.orders}`,Actual:orderRows.length,Status:orderRows.length?'IMPORTED':'NO ROWS',Source:orderSource?.file||''},
      {Check:'Data Shipping source rows',Rule:`Filter by Manifest Scan Date ${r.reportDates.data}`,Actual:dataRows.length,Status:dataRows.length?'IMPORTED':'NO ROWS',Source:r.selectedFc?.file||''},
      {Check:'Shipping Status period',Rule:`${r.reportDates.statusStart} through ${r.reportDates.statusEnd}; distinct Work Orders only`,Actual:computedDates.join(', ')||'No FC rows',Status:'NOT CERTIFIED — unit unconfirmed',Source:orderSource?.file||''},
      {Check:'UPS source events',Rule:'Retain every UPS source event',Actual:upsRaw.length,Status:upsRaw.length?'IMPORTED':'MISSING',Source:state.sheets.find(s=>s.kind==='UPS')?.file||''},
      {Check:'UPS distinct Tracking Numbers',Rule:'Group by Tracking Number for summary only; not shoe-pair count',Actual:new Set(upsRaw.map(x=>norm(value(x,'Tracking Number'))).filter(Boolean)).size,Status:'CALCULATED',Source:'UPS Tracking Number'},
      {Check:'SF manifest detail rows',Rule:'Keep SF Baozun reconciliation separate; quantity differences against MB52 are reference-only',Actual:sfExport.length,Status:sfExport.length?'IMPORTED':'MISSING',Source:state.sheets.find(s=>s.kind==='SF Baozun')?.file||''},
      {Check:'Manual output row comparison',Rule:'Compare uploaded manual outputs field by field',Actual:r.manual.lines.length,Status:r.manual.lines.length?'REVIEW':'NO BASELINE',Source:'See Manual Field Differences worksheet'},
      {Check:'Expect Shipment',Rule:'Business rule must be confirmed before calculation',Actual:'Not calculated',Status:'PENDING RULE',Source:'No estimate is generated'},
      {Check:'Shipping Status history',Rule:'Preserve uploaded Summary and Detail shipment history',Actual:historySummary.length?'History loaded':'No history file imported',Status:historySummary.length?'IMPORTED':'MISSING',Source:state.sheets.find(s=>s.kind==='Shipping Status Summary')?.file||''},
      {Check:'Actually Shipment unit',Rule:'Unique Work Order count is not certified as shoe-pair quantity',Actual:'Distinct Work Orders',Status:'NOT CERTIFIED',Source:'Business rule pending'},
      {Check:'Golden Sample certification',Rule:'Independent raw FC, SAP, carrier files and manual outputs are required',Actual:'Not verified',Status:'PENDING',Source:'No independent sample in this session'}
    ];
    const files={};
    const make=(name, sheets)=>{const wb=XLSX.utils.book_new();for(const [sheetName,rows] of sheets){const data=rows?.length?rows:[{'Status':'No rows','Note':'No matching source rows were imported'}];const ws=XLSX.utils.json_to_sheet(data,{header:headersOf(data)});ws['!freeze']={xSplit:0,ySplit:1};if(ws['!ref'])ws['!autofilter']={ref:ws['!ref']};XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));}files[name]=XLSX.write(wb,{bookType:'xlsx',type:'array',cellStyles:true});};
    if(orderRows.length)make(outputNames[0],[['Orders Update',orderRows]]);
    if(dataRows.length&&r.counts.sap>0&&r.counts.mb52>0)make(outputNames[1],[['Data Shipping',dataRows]]);
    if(computedDates.length)make(outputNames[2],[['Summary',statusSummary],['Detail shipment',statusDetail]]);
    if(upsRaw.length)make(outputNames[3],[['UPS Reconciliation',upsExport],['Tracking Summary',englishRows(r.upsRows)]]);
    if(sfExport.length)make(outputNames[4],[['SF Baozun Reconciliation',sfExport]]);
    if(r.counts.manualFiles>0)make(outputNames[6],[['Golden Sample Differences',goldenDifferences],['Comparison Summary',englishRows(r.manual.summaries.map(({details,...x})=>x))]]);
    r.templateStatus=[];
    for(const [type,name,plans] of [
      ['orders',outputNames[0],[{kind:'fc',rows:orderRows}]],
      ['data',outputNames[1],[{kind:'fc',rows:dataRows},{kind:'sap',rows:state.sheets.filter(s=>s.kind==='SAP ZRSD1019'&&(!r.sourceDecisions?.['SAP ZRSD1019']?.selected||s.file===r.sourceDecisions['SAP ZRSD1019'].selected)).flatMap(s=>s.rows)},{kind:'stock',rows:state.sheets.filter(s=>s.kind==='SAP MB52'&&(!r.sourceDecisions?.['SAP MB52']?.selected||s.file===r.sourceDecisions['SAP MB52'].selected)).flatMap(s=>s.rows)},{kind:'awb',rows:state.sheets.filter(s=>s.kind==='AWB Map'&&(!r.sourceDecisions?.['AWB Map']?.selected||s.file===r.sourceDecisions['AWB Map'].selected)).flatMap(s=>s.rows)}]],
      ['status',outputNames[2],[]]
    ]){
      const template=state.templates[type];
      if(!files[name]){r.templateStatus.push(name+'：缺少必要來源或沒有符合資料列');continue;}
      if(template){try{if(type==='status'){let bytes=template.bytes;for(const date of computedDates){const verifiedPairs=pairCount(totalDateOrders.get(date),date);const products=[...productDateOrders.values()].filter(x=>x.date===date).map(x=>({product:x.product,count:x.workOrders.size}));bytes=await NBYTemplates.fill(bytes,plans,type,{date,actual:totalDateOrders.get(date).size,products,expectDate:verifiedPairs===null?null:addDay(date),expect:verifiedPairs});}if(!computedDates.length)throw Error('所選期間沒有 FC 出貨列可更新');files[name]=bytes;const updated={...template,bytes:new Uint8Array(bytes)};try{saveTemplate(type,updated);state.templates[type]=updated;}catch(error){console.warn('Shipping history could not be saved',error);r.templateStatus.push(name+'：未能儲存歷史至瀏覽器');}}else files[name]=await NBYTemplates.fill(template.bytes,plans,type,{});r.templateStatus.push(name+'：原樣板格式');}catch(e){delete files[name];r.templateStatus.push(name+'：未輸出，'+e.message);}}
      else {const draft='DRAFT '+name;files[draft]=files[name];delete files[name];r.templateStatus.push(name+'：未載入樣板，僅提供 DRAFT 資料版');}
    }
    const reportState=(name,ready)=>{const official=!!files[name],draft=!!files['DRAFT '+name];return !ready?'Missing Required Source':official?'Generated – Pending Validation':draft?'Generated – Draft Pending Validation':'Failed'};
    r.reportStatus=[
      {name:'Orders Update',state:reportState(outputNames[0],orderRows.length>0)},
      {name:'Data Shipping',state:reportState(outputNames[1],dataRows.length>0&&r.counts.sap>0&&r.counts.mb52>0)},
      {name:'Shipping Status',state:reportState(outputNames[2],computedDates.length>0)},
      {name:'UPS Reconciliation',state:reportState(outputNames[3],upsRaw.length>0)},
      {name:'SF Baozun Reconciliation',state:reportState(outputNames[4],sfExport.length>0&&dataRows.length>0&&r.counts.sap>0&&r.counts.mb52>0)},
      {name:'Exception Report',state:'Generated – Pending Review'},
      {name:'Golden Sample Difference Report',state:reportState(outputNames[6],r.counts.manualFiles>0)},
      {name:'Golden Sample Certification',state:'Pending – Independent Raw Source Required'}
    ];
    testSummary.push(...r.reportStatus.map(x=>({Check:`Report: ${x.name}`,Rule:'Output generation is separate from manual verification',Actual:x.state,Status:x.state,Source:'Processing result'})));
    testSummary.push({Check:'Output validation',Rule:'Generated workbook and source readiness still require manual row-by-row review',Actual:Object.keys(files).join(', '),Status:'PENDING MANUAL VALIDATION',Source:'Generated file list'});
    make(outputNames[5],[['Exceptions',exceptions],['SF Quantity References',englishRows(r.sfRows)],['UPS Quantity Status',englishRows(r.upsRows)],['Shipping Status Diagnostics',englishRows(statusSummary)],['Test Result',testSummary],['Template Status',r.templateStatus.map(Status=>({Status}))],['Processing Log',state.files.map(f=>({File:f.name,Bytes:f.size,ImportedAt:f.importedAt,Types:f.types.join(', '),Error:f.error||'',Validation:JSON.stringify(f.validation||[]),ReportDates:JSON.stringify(r.reportDates)}))],['Golden Sample Differences',goldenDifferences],['Golden Sample Summary',englishRows(r.manual.summaries.map(({details,...x})=>x))]]);
    await buildReady;
    const audit={build:{...window.NBYBuild},version:'2.1.5',sapPaaByDate:Object.fromEntries(computedDates.map(date=>[date,pairCount(totalDateOrders.get(date),date)])),processedAt:new Date().toISOString(),manifestDate:r.target,reportDates:r.reportDates,sourceSelections:r.sourceDecisions,source:state.files.map(f=>({name:f.name,size:f.size,importedAt:f.importedAt,types:f.types,validation:f.validation||[],error:f.error||''})),selectedFc:{file:r.selectedFc.file,sheet:r.selectedFc.sheet},counts:r.counts,exceptions:r.anomalies.length,templates:r.templateStatus,reportStatus:r.reportStatus,expectShipment:'PREVIOUS_MANIFEST_PAA_VERIFIED',goldenSample:'NOT CERTIFIED'};
    r.audit=audit;try{const history=JSON.parse(localStorage.getItem('nby-audit')||'[]');history.unshift(audit);localStorage.setItem('nby-audit',JSON.stringify(history.slice(0,30)));}catch(e){};
    r.outputFiles=files;r.outputTestSummary=testSummary;renderReportStatuses();renderOutputFiles();return files;
  }
  function downloadBytes(bytes,name){const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
  function renderReportStatuses(){
    const root=$('reportStatuses');if(!root)return;
    const nameKeys={'Orders Update':'reportOrders','Data Shipping':'reportData','Shipping Status':'reportStatus','UPS Reconciliation':'reportUps','SF Baozun Reconciliation':'reportSf','Exception Report':'reportException','Golden Sample Difference Report':'reportGoldenDiff','Golden Sample Certification':'reportGoldenCert'};
    const stateKeys={'Missing Required Source':'reportMissing','Generated – Pending Validation':'reportGenerated','Generated – Draft Pending Validation':'reportDraft','Failed':'reportFailed','Generated – Pending Review':'reportReview','Pending – Independent Raw Source Required':'reportGolden'};
    root.innerHTML=(state.result.reportStatus||[]).map(x=>`<div class="report-status"><strong>${escapeHtml(t(nameKeys[x.name]||x.name))}</strong><span>${escapeHtml(t(stateKeys[x.state]||x.state))}</span></div>`).join('');
  }
  function renderOutputFiles(){
    const root=$('outputFiles');if(!root)return;
    root.innerHTML=Object.keys(state.result.outputFiles).map(name=>`<button class="output-file" data-output="${escapeHtml(name)}">↓ ${escapeHtml(t('download'))} ${escapeHtml(name)}</button>`).join('')+`<span class="output-note">${escapeHtml(t('outputNote'))}</span>`;
    root.querySelectorAll('[data-output]').forEach(btn=>btn.addEventListener('click',()=>{const bytes=state.result?.outputFiles?.[btn.dataset.output];if(bytes)downloadBytes(bytes,btn.dataset.output);}));
  }
  async function exportXlsx(){
    const files=state.result?.outputFiles||await buildOutputs();
    if(!window.JSZip){status('zipUnavailable');return;}
    try{const zip=new JSZip();for(const [name,bytes] of Object.entries(files))zip.file(name,bytes);zip.file('Processing Audit.json',JSON.stringify(state.result.audit,null,2));const blob=await zip.generateAsync({type:'blob'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=`NBY_Shipping_Outputs_${state.result.target}.zip`;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}catch(e){console.error(e);status('zipUnavailable');}
  }
  Object.assign(i18n.zh,{xlsxTemplateRequired:'已辨識人工成品。請在 Excel 另存為 XLSX，再匯入以設定樣板。',templateSaveFailed:'樣板無法儲存，請檢查瀏覽器儲存空間。',dateLabel:'Report Date（報表日期）',missingPreviousFc:'缺少 {date} 的 Nike FC 原始資料；Orders Update 必須使用 Report Date 前一天的 Manifest Scan Date。請批次補入該日期原始檔。',fcSourceLabel:'已匯入 FC 日期來源（系統自動分配）'});
  Object.assign(i18n.en,{xlsxTemplateRequired:'Manual report detected. Save it as XLSX in Excel and import it to set the template.',templateSaveFailed:'Could not save template. Check browser storage.',dateLabel:'Report Date',missingPreviousFc:'Missing raw Nike FC data for {date}. Orders Update requires the previous day’s Manifest Scan Date. Import the raw file for that date.',fcSourceLabel:'Imported FC sources (automatically assigned)'});
  Object.assign(i18n.vi,{xlsxTemplateRequired:'Đã nhận diện báo cáo thủ công. Lưu thành XLSX trong Excel rồi tải lên để thiết lập mẫu.',templateSaveFailed:'Không thể lưu mẫu. Vui lòng kiểm tra dung lượng trình duyệt.',dateLabel:'Ngày báo cáo (Report Date)',missingPreviousFc:'Thiếu dữ liệu Nike FC gốc ngày {date}. Orders Update sử dụng Manifest Scan Date của ngày trước ngày báo cáo. Vui lòng tải thêm tệp gốc của ngày đó.',fcSourceLabel:'Nguồn FC đã tải (tự động phân bổ)'});
  window.NBYBuild={version:'2.1.5',buildId:'shipping-2.1.5-r1',gitCommitSha:null,deployedAt:null};
  const buildLabel=document.getElementById('buildInfo');
  function showBuild(){
    if(!buildLabel)return;
    const b=window.NBYBuild;
    buildLabel.textContent=`Version ${b.version} · Build ${b.buildId} · Git SHA ${b.gitCommitSha||'—'} · Deployment UTC ${b.deployedAt||'—'}`;
  }
  showBuild();
  const buildReady=typeof fetch==='function'?fetch('build-info.json',{cache:'no-store'}).then(r=>r.ok?r.json():null).then(info=>{
    if(info?.buildId===window.NBYBuild.buildId){Object.assign(window.NBYBuild,info);showBuild();}
  }).catch(()=>{}):Promise.resolve();
  const templateKey=type=>'nby-template-v1-'+type;
  function saveTemplate(type, template) {
    let binary='';
    for(let offset=0;offset<template.bytes.length;offset+=16384)
      binary+=String.fromCharCode(...template.bytes.subarray(offset,offset+16384));
    localStorage.setItem(templateKey(type),JSON.stringify({name:template.name,bytes:btoa(binary)}));
  }
  function restoreTemplates() {
    for(const type of ['orders','data','status']) {
      try {
        const saved=JSON.parse(localStorage.getItem(templateKey(type))||'null');
        if(saved)state.templates[type]={name:saved.name,bytes:Uint8Array.from(atob(saved.bytes),x=>x.charCodeAt(0))};
      } catch(error) { console.warn('Template could not be restored',type,error); }
    }
  }
  function renderTemplateNames(){
    document.querySelectorAll('[data-template]').forEach(input=>{
      const template=state.templates[input.dataset.template];
      input.title=template?.name||'';
      input.parentElement.querySelector('small')?.remove();
      if(template){const label=document.createElement('small');label.textContent=template.name;input.after(label);}
    });
  }
    function applyLanguage(next){
    language=next;renderTemplateNames();document.documentElement.lang=language==='zh'?'zh-Hant':language==='vi'?'vi':'en';
    document.querySelectorAll('[data-i18n]').forEach(el=>{el.textContent=t(el.dataset.i18n);});
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el=>{el.placeholder=t(el.dataset.i18nPlaceholder);});
    const nextLang=languageOrder[(languageOrder.indexOf(language)+1)%languageOrder.length];
    $('langToggle').textContent=nextLang==='zh'?'繁中':nextLang.toUpperCase();$('langToggle').setAttribute('aria-label',t('switchLanguage'));
    renderSources();renderSourceChoices();renderSourceOverrides();
    $('fcSource').disabled=true;const advanced=$('ordersDate').closest('details');if(advanced)advanced.hidden=true;
    if(state.result){renderResults();renderReportStatuses();renderOutputFiles();}
    if(state.statusMessage)status(state.statusMessage.key,state.statusMessage.params||{});
  }
  function invalidate(){state.result=null;$('results').classList.add('hidden');status('dataChanged');}
  $('asOfDate').addEventListener('change',invalidate);$('fcSource').addEventListener('change',invalidate);$('sourceOverrides').addEventListener('change',e=>{const kind=e.target.dataset.sourceOverride;if(!kind)return;state.sourceOverrides[kind]=e.target.value;invalidate();renderSourceOverrides();});['ordersDate','dataDate','statusStart','statusEnd'].forEach(id=>$(id).addEventListener('change',invalidate));
  document.querySelectorAll('[data-template]').forEach(input=>input.addEventListener('change',async()=>{invalidate();const f=input.files[0];if(!f)return;state.templates[input.dataset.template]={bytes:new Uint8Array(await f.arrayBuffer()),name:f.name};try{saveTemplate(input.dataset.template,state.templates[input.dataset.template]);}catch(error){console.warn(error);$('runStatus').textContent=error.message;renderTemplateNames();return;}renderTemplateNames();status('templateLoaded',{name:f.name});}));
  const clearReplacement=kind=>{const old=state.files.filter(f=>f.expectedKind===kind||f.types?.includes(kind));const names=new Set(old.map(f=>f.name));state.files=state.files.filter(f=>!names.has(f.name));state.sheets=state.sheets.filter(s=>!names.has(s.file));state.replacingKind='';renderSources();renderSourceChoices();};
  document.querySelectorAll('[data-source-file]').forEach(input=>input.addEventListener('change',async e=>{const kind=input.dataset.sourceFile;if(state.replacingKind===kind)clearReplacement(kind);await ingest(e.target.files,kind);input.value='';}));
  $('bulkFiles').addEventListener('change',async e=>{const files=[...e.target.files];$('bulkSelection').textContent=files.map(f=>f.name).join(' · ');await ingest(files);e.target.value='';});
  const drop=$('bulkDropZone');drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('dragover');});drop.addEventListener('dragleave',()=>drop.classList.remove('dragover'));drop.addEventListener('drop',async e=>{e.preventDefault();drop.classList.remove('dragover');const files=[...e.dataTransfer.files];$('bulkSelection').textContent=files.map(f=>f.name).join(' · ');await ingest(files);});
  document.querySelectorAll('[data-replace-source]').forEach(button=>button.addEventListener('click',()=>{state.replacingKind=button.dataset.replaceSource;document.querySelector(`[data-source-file="${button.dataset.replaceSource}"]`)?.click();}));
  $('runButton').addEventListener('click',async()=>{if(state.busy)return;state.busy=true;$('runButton').disabled=true;invalidate();status('runProcessing');try{await new Promise(resolve=>setTimeout(resolve,0));await calculate();status('runDone');}catch(e){console.error(e);$('runStatus').textContent=`${t('runFailed')}: ${uiText(e.message)}`;}finally{state.busy=false;$('runButton').disabled=false;}});
  $('clearButton').addEventListener('click',()=>{state.replacingKind='';document.querySelectorAll('[data-template]').forEach(e=>e.value='');state.files=[];state.sheets=[];state.result=null;state.sourceOverrides={};document.querySelectorAll('[data-source-file]').forEach(input=>input.value='');$('bulkFiles').value='';$('bulkSelection').textContent='';$('results').classList.add('hidden');$('runStatus').textContent='';$('runButton').disabled=true;$('clearButton').disabled=true;renderSources();renderSourceChoices();renderSourceOverrides();});
  $('exportButton').addEventListener('click',exportXlsx);$('tableSearch').addEventListener('input',renderTable);
  document.querySelectorAll('.tab').forEach(tab=>tab.addEventListener('click',()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));tab.classList.add('active');state.activeTab=tab.dataset.tab;renderTable();}));
  ['ordersDate','dataDate','statusStart','statusEnd'].forEach(id=>$(id).value='');
  $('langToggle').addEventListener('click',()=>applyLanguage(languageOrder[(languageOrder.indexOf(language)+1)%languageOrder.length]));
  restoreTemplates();
  applyLanguage('zh');
})();
