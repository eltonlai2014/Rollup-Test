const _ = require('lodash');
const log4js = require('log4js');
const logger = log4js.getLogger('ReportExcel');

const Excel = require('exceljs');
const SiteChect = require('./chart/SiteChart');

class ReportExcel {

    constructor(options) {
        this.options = options;
    }

    getParamValue(paramName, defaultValue) {
        if (this.options[paramName] != undefined) {
            return this.options[paramName];
        }
        return defaultValue;
    }
    genReport(report) {
        const fillCell = (aRow, dataIndex, col, value, style_even, style) =>{
            let aCell = aRow.getCell(col);
            if (dataIndex % 2 == 0) {
                aCell.style = style_even;
            } else {
                aCell.style = style;
            }
            if(value == undefined){
                value = '';
            }            
            aCell.value = value;
            return aCell;
        }

        const fillTitleCell = (aRow, col, value, style) => {
            let aCell = aRow.getCell(col);
            aCell.style = style;
            aCell.value = value;
            return aCell;
        }

        const buildRankMap = (data) => {
            let rMap = {};
            for (let i = 0; i < data.RankList.length; i++) {
                let aEventType = data.RankList[i].EventType;
                let aRankList = data.RankList[i].Rank;
                rMap[aEventType] = aRankList;
            }
            // console.log('rMap', rMap);
            return rMap;       
        }

        const getRankValue = (rMap, eventType, rankOrder) => {
            let aRankList = rMap[eventType];
            if(aRankList == null){
                aRankList = [];
            }
            if(aRankList[rankOrder] != undefined){
                return aRankList[rankOrder];
            }
            return { SiteId: '', SiteName: ''};
        }


        return new Promise(async (resolve, reject) => {
            this.report = report;
            logger.info('ReportExcel init', this.options);
            // 依據選擇區塊計算白底列數
            let totalRows = 11;                                                         // 11 title block
            let OperationBlock = this.getParamValue('OperationBlock', true);            // 19
            let NetworkDeviceBlock = this.getParamValue('NetworkDeviceBlock', true);    // 19
            let CriticalEventBlock = this.getParamValue('CriticalEventBlock', true);    // 41
            let WarningEventBlock = this.getParamValue('WarningEventBlock', true);      // 41
            if(OperationBlock) {
                // 這裡需要看有幾個月的資料
                totalRows+= (16 + report.reportData.data.length);
            }
            if(NetworkDeviceBlock) {
                // 這裡需要看有幾個月的資料
                totalRows+= (16 + report.reportData2.data.length);
            }
            if(CriticalEventBlock) {
                totalRows+= 41;
            }
            if(WarningEventBlock) {
                totalRows+= 41;
            }

            // 讀取樣板檔案
            this.workbook = new Excel.Workbook();
            await this.workbook.xlsx.readFile(this.options.StyleFile);
            const worksheet = this.workbook.getWorksheet(this.options.StyleSheet);

            // 定義事件分類
            const ETYPE_Total = 99 ;
            const ETYPE_Device_Unreachable = 1 ;
            const ETYPE_Network_Alert = 2 ;
            const ETYPE_Ethernet_Port_Alert = 3 ;
            const ETYPE_Fiber_Port_Alert = 4 ;
            const ETYPE_Power_Supply_Alert = 5 ;
            const ETYPE_Network_Intrusion_Alert = 6 ;
            const ETYPE_Device_Security_Alert = 7 ;
            const ETYPE_Device_Status_Alert = 8 ;
            const ETYPE_MXview_One_Server_Alert = 9 ;
            const ETYPE_GOOSE = 10 ;
            // [繪圖物件準備]
            const aSiteChect = new SiteChect(this.options.chartOptions.chartWidth || 540, this.options.chartOptions.chartHeight || 300, this.options.chartOptions);

            // 取得欄位寬度
            let colWidth = [];
            for (let i = 1; i <= 20; i++) {
                let aColumn = worksheet.getColumn(i);
                colWidth.push(aColumn.width);
            }
            console.log('colWidth', colWidth);
            // 取得列高
            // 列高目前都一樣，依據字型大小縮放

            // 取得報表基本樣式
            const bgGray = worksheet.getCell('B1').style;
            const bgWhite = worksheet.getCell('B2').style;

            const reportTitle = worksheet.getCell('H3').style;
            const reportSummary = worksheet.getCell('C7').style;
            const blockTitle = worksheet.getCell('C12').style;
            const blockContent = worksheet.getCell('F12').style;
            const tableTitle_L = worksheet.getCell('C23').style;
            const tableBody_L = worksheet.getCell('C24').style;
            const tableTitle_R = worksheet.getCell('D23').style;
            const tableBody_R = worksheet.getCell('D24').style;

            const tableTitle2_L = worksheet.getCell('C52').style;
            const tableTitle2_R = worksheet.getCell('D52').style;
            const tableBody2_L = worksheet.getCell('C53').style;
            const tableBody2_R = worksheet.getCell('D53').style;
            const tableBody2_L_even = worksheet.getCell('C54').style;
            const tableBody2_R_even = worksheet.getCell('D54').style;

            const ServiceSummaryRows = 7;

            // 產生新的活頁表
            let outputSheet = this.workbook.addWorksheet("台電測試站", {
                pageSetup: { paperSize: 12 }
            });

            // 設定欄位寬度
            for (let i = 1; i <= 20; i++) {
                let aColumn = outputSheet.getColumn(i);
                aColumn.width = colWidth[i - 1];
            }

            const BG_ROW_COUNT = 200;
            const RP_ROW_COUNT = totalRows + 1;
            // RP_ROW_COUNT 會依據選擇報表區塊而有不同

            // 填報表背景顏色
            for (let i = 1; i <= BG_ROW_COUNT; i++) {
                let aRow = outputSheet.getRow(i);
                // 灰底
                for (let j = 1; j <= 20; j++) {
                    let aCell = aRow.getCell(j);
                    aCell.style = bgGray;
                }
                // 白底
                if (i > 1 && i < RP_ROW_COUNT) {
                    for (let j = 2; j <= 14; j++) {
                        let aCell = aRow.getCell(j);
                        aCell.style = bgWhite;
                    }
                }
            }

            // [報表標題]
            let startRow = 3;
            let aRow = outputSheet.getRow(startRow);
            let aCell = aRow.getCell(8);
            aCell.value = '公司報表_台電';
            aCell.style = reportTitle;
            aCell.font = reportTitle.font;
            startRow += 2;

            // ====================================================================================================
            // 報表Summary
            // ====================================================================================================
            const SummaryRows = 5;
            for (let i = startRow; i < startRow + SummaryRows; i++) {
                let aRow = outputSheet.getRow(i);
                // 填滿區塊底色
                for (let i = 3; i <= 13; i++) {
                    let aCell = aRow.getCell(i);
                    aCell.style = reportSummary;
                }
                switch (i) {
                    case startRow + 1: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  公司　${report.companyName}`;
                        break;
                    }
                    case startRow + 2: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  報表產生時間　${report.reportDate}`;
                        break;
                    }
                    case startRow + 3: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  報告內容時間　${report.reportSDate}~${report.reportEate}`;
                        break;
                    }
                }
            }
            startRow += SummaryRows;
            startRow += 2;

            // ====================================================================================================
            // 站台營運服務狀態
            // ====================================================================================================
            if(OperationBlock) {
                aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:M${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = '站台營運服務狀態';
                aCell.style = blockTitle;
                const imgStartRow = startRow + 1;
                startRow += 2;
                // [服務狀態說明]
                const expressLabel = [`  各站台 MXview 主機營運狀況`, `  以是否發生 MXview One Server Alert 區分`, `  Critical Site  當月發生過 Critical Event`,
                    `  Warning Site  當月未發生Critical Event 但有 Warning Event`, `  Health Site  當月未發生 Critical Event 和 Warning Event`,
                ];
                for (let i = startRow; i < startRow + ServiceSummaryRows; i++) {
                    let aRow = outputSheet.getRow(i);
                    // 填滿區塊底色
                    for (let j = 3; j <= 7; j++) {
                        let aCell = aRow.getCell(j);
                        aCell.style = reportSummary;
                    }

                    // 填欄位說明
                    switch (i) {
                        case startRow + 1:
                        case startRow + 2:
                        case startRow + 3:
                        case startRow + 4:
                        case startRow + 5:
                            outputSheet.mergeCells(`C${i}:G${i}`);
                            let aCell = aRow.getCell(3);
                            aCell.value = expressLabel[i - startRow - 1];
                            break;
                    }
                }
                startRow += ServiceSummaryRows;
                startRow += 2;

                // [彙總數字]
                let title = ['', 'Health Site', 'Warning Site', 'Critical Site', 'Total'];
                for (let i = 0; i < report.reportData.data.length; i++) {
                    let aInfo = report.reportData.data[i];
                    aInfo.Total = aInfo.Health + aInfo.Warning + aInfo.Critical;
                }
                // 將資料依照日期排序
                report.reportData.data = _.sortBy(report.reportData.data, ['Month']);

                for (let i = startRow; i <= startRow + (report.reportData.data.length + 1); i++) {
                    let aRow = outputSheet.getRow(i);
                    switch (i) {
                        case startRow: {
                            // 表格標題
                            for (let j = 0; j < 5; j++) {
                                let aCell = aRow.getCell(j + 3);
                                aCell.value = title[j];
                                aCell.style = tableTitle_R;
                            }
                            break;
                        }
                        default:
                            // 數值
                            let aInfo = report.reportData.data[i - startRow - 1];
                            if (aInfo) {
                                let aCell = aRow.getCell(3);
                                aCell.style = tableBody_L;
                                aCell.value = aInfo.Month;
                                aCell = aRow.getCell(4);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Health;
                                aCell = aRow.getCell(5);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Warning;
                                aCell = aRow.getCell(6);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Critical;
                                aCell = aRow.getCell(7);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Total;
                            }

                            break;
                    }
                }

                // [繪圖]
                // 繪製圖檔並產生png buffer
                aSiteChect.setChartData(report.reportData, SiteChect.Type.COMPANY).paint();
                const buffer = aSiteChect.getCanvasBuffer();

                // 加入excel workbook
                const imageId1 = this.workbook.addImage({
                    buffer: buffer,
                    extension: 'png',
                });
                outputSheet.addImage(imageId1, { tl: { col: 7.5, row: imgStartRow }, ext: { width: this.options.chartOptions.chartWidth, height: this.options.chartOptions.chartHeight } });

                startRow += report.reportData.data.length;
                startRow += 5;
            }

            // ====================================================================================================
            // 站台網路及設備監控狀態
            // ====================================================================================================
            if(NetworkDeviceBlock) {
                aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:M${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = '站台網路及設備監控狀態';
                aCell.style = blockTitle;
                const imgStartRow = startRow + 1;
                startRow += 2;

                // [服務狀態說明]
                const SiteSummaryRows = 7;
                const expressLabel_2 = [`  各站台網路及設備狀況`, `  以是否發生 MXview One Server Alert 之外的 Alert 區分`, `  Critical Site  當月發生過 Critical Event`,
                    `  Warning Site  當月未發生Critical Event 但有 Warning Event`, `  Health Site  當月未發生 Critical Event 和 Warning Event`,
                ];
                for (let i = startRow; i < startRow + SiteSummaryRows; i++) {
                    let aRow = outputSheet.getRow(i);
                    // 填滿區塊底色
                    for (let j = 3; j <= 7; j++) {
                        let aCell = aRow.getCell(j);
                        aCell.style = reportSummary;
                    }

                    // 填欄位說明
                    switch (i) {
                        case startRow + 1:
                        case startRow + 2:
                        case startRow + 3:
                        case startRow + 4:
                        case startRow + 5:
                            outputSheet.mergeCells(`C${i}:G${i}`);
                            let aCell = aRow.getCell(3);
                            aCell.value = expressLabel_2[i - startRow - 1];
                            break;
                    }
                }
                startRow += ServiceSummaryRows;
                startRow += 2;

                // [彙總數字]
                let title_2 = ['', 'Health Site', 'Warning Site', 'Critical Site', 'Total'];
                for (let i = 0; i < report.reportData2.data.length; i++) {
                    let aInfo = report.reportData2.data[i];
                    aInfo.Total = aInfo.Health + aInfo.Warning + aInfo.Critical;
                }
                // 將資料依照日期排序
                report.reportData2.data = _.sortBy(report.reportData2.data, ['Month']);

                for (let i = startRow; i <= startRow + (report.reportData2.data.length + 1); i++) {
                    let aRow = outputSheet.getRow(i);
                    switch (i) {
                        case startRow: {
                            // 表格標題
                            for (let j = 0; j < 5; j++) {
                                let aCell = aRow.getCell(j + 3);
                                aCell.value = title_2[j];
                                aCell.style = tableTitle_R;
                            }
                            break;
                        }
                        default:
                            // 數值
                            let aInfo = report.reportData2.data[i - startRow - 1];
                            if (aInfo) {
                                let aCell = aRow.getCell(3);
                                aCell.style = tableBody_L;
                                aCell.value = aInfo.Month;
                                aCell = aRow.getCell(4);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Health;
                                aCell = aRow.getCell(5);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Warning;
                                aCell = aRow.getCell(6);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Critical;
                                aCell = aRow.getCell(7);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Total;
                            }

                            break;
                    }
                }

                // [繪圖]
                // 繪製圖檔並產生png buffer
                aSiteChect.setChartData(report.reportData2, SiteChect.Type.COMPANY).paint();
                const buffer2 = aSiteChect.getCanvasBuffer();

                // 加入excel workbook
                const imageId2 = this.workbook.addImage({
                    buffer: buffer2,
                    extension: 'png',
                });
                outputSheet.addImage(imageId2, { tl: { col: 7.5, row: imgStartRow }, ext: { width: this.options.chartOptions.chartWidth, height: this.options.chartOptions.chartHeight } });
                startRow += report.reportData2.data.length;
                startRow += 5;
            }

            // ====================================================================================================
            // Critical Event 發生次數 Top 10
            // ====================================================================================================
            if(CriticalEventBlock) {
                let rMap = buildRankMap(report.reportData3);
                aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:M${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = 'Critical Event 發生次數 Top 10';
                aCell.style = blockTitle;
                startRow += 2;

                // [Type 1~3]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 4, 'Total', tableTitle2_R);
                        fillTitleCell(aRow, 6, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 7, 'Device unreachable', tableTitle2_R);

                        fillTitleCell(aRow, 9, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 10, 'Network alert', tableTitle2_R);
                        fillTitleCell(aRow, 12, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 13, 'Ethernet port alert', tableTitle2_R);

                    } else {
                        let rankInfo = getRankValue(rMap, ETYPE_Total, i - startRow - 1);
                        fillCell(aRow, i - startRow, 3, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 4, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Device_Unreachable, i - startRow - 1);
                        fillCell(aRow, i - startRow, 6, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 7, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Network_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 9, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 10, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Ethernet_Port_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 12, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 13, rankInfo.Times, tableBody2_R_even, tableBody2_R);
                    }
                }
                startRow += 13;
                // [Type 4~7]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 4, 'Fiber port alert', tableTitle2_R);
                        fillTitleCell(aRow, 6, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 7, 'Power supply alert', tableTitle2_R);

                        fillTitleCell(aRow, 9, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 10, 'Network intrusion alert', tableTitle2_R);
                        fillTitleCell(aRow, 12, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 13, 'Device security alert', tableTitle2_R);

                    } else {
                        let rankInfo = getRankValue(rMap, ETYPE_Fiber_Port_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 3, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 4, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Power_Supply_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 6, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 7, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Network_Intrusion_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 9, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 10, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Device_Security_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 12, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 13, rankInfo.Times, tableBody2_R_even, tableBody2_R);
                    }
                }
                startRow += 13;
                // [Type 8~10]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 4, 'Device Status Alert', tableTitle2_R);
                        fillTitleCell(aRow, 6, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 7, 'MXview One Server Alert', tableTitle2_R);
                        fillTitleCell(aRow, 9, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 10, 'GOOSE', tableTitle2_R);
                    } else {
                        let rankInfo = getRankValue(rMap, ETYPE_Device_Status_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 3, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 4, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_MXview_One_Server_Alert, i - startRow - 1);
                        fillCell(aRow, i - startRow, 6, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 7, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_GOOSE, i - startRow - 1);
                        fillCell(aRow, i - startRow, 9, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 10, rankInfo.Times, tableBody2_R_even, tableBody2_R);
                    }
                }
                startRow += 13;
            }

            // ====================================================================================================
            // Warning Event 發生次數 Top 10
            // ====================================================================================================
            if(WarningEventBlock) {
                let rMap = buildRankMap(report.reportData4);
                aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:M${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = 'Warning Event 發生次數 Top 10';
                aCell.style = blockTitle;
                startRow += 2;

                // [Type 1~3]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 4, 'Total', tableTitle2_R);
                        fillTitleCell(aRow, 6, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 7, 'Device unreachable', tableTitle2_R);

                        fillTitleCell(aRow, 9, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 10, 'Network alert', tableTitle2_R);
                        fillTitleCell(aRow, 12, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 13, 'Ethernet port alert', tableTitle2_R);

                    } else {
                        let rankInfo = getRankValue(rMap, ETYPE_Total, i - startRow - 1);
                        fillCell(aRow, i, 3, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 4, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Device_Unreachable, i - startRow - 1);
                        fillCell(aRow, i, 6, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 7, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Network_Alert, i - startRow - 1);
                        fillCell(aRow, i, 9, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 10, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Ethernet_Port_Alert, i - startRow - 1);
                        fillCell(aRow, i, 12, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 13, rankInfo.Times, tableBody2_R_even, tableBody2_R);
                    }
                }
                startRow += 13;
                // [Type 4~7]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 4, 'Fiber port alert', tableTitle2_R);
                        fillTitleCell(aRow, 6, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 7, 'Power supply alert', tableTitle2_R);

                        fillTitleCell(aRow, 9, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 10, 'Network intrusion alert', tableTitle2_R);
                        fillTitleCell(aRow, 12, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 13, 'Device security alert', tableTitle2_R);

                    } else {
                        let rankInfo = getRankValue(rMap, ETYPE_Fiber_Port_Alert, i - startRow - 1);
                        fillCell(aRow, i, 3, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 4, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Power_Supply_Alert, i - startRow - 1);
                        fillCell(aRow, i, 6, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 7, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Network_Intrusion_Alert, i - startRow - 1);
                        fillCell(aRow, i, 9, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 10, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_Device_Security_Alert, i - startRow - 1);
                        fillCell(aRow, i, 12, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 13, rankInfo.Times, tableBody2_R_even, tableBody2_R);
                    }
                }
                startRow += 13;
                // [Type 8~10]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 4, 'Device Status Alert', tableTitle2_R);
                        fillTitleCell(aRow, 6, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 7, 'MXview One Server Alert', tableTitle2_R);
                        fillTitleCell(aRow, 9, 'Site', tableTitle2_L);
                        fillTitleCell(aRow, 10, 'GOOSE', tableTitle2_R);
                    } else {
                        let rankInfo = getRankValue(rMap, ETYPE_Device_Status_Alert, i - startRow - 1);
                        fillCell(aRow, i, 3, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 4, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_MXview_One_Server_Alert, i - startRow - 1);
                        fillCell(aRow, i, 6, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 7, rankInfo.Times, tableBody2_R_even, tableBody2_R);

                        rankInfo = getRankValue(rMap, ETYPE_GOOSE, i - startRow - 1);
                        fillCell(aRow, i, 9, rankInfo.SiteName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i, 10, rankInfo.Times, tableBody2_R_even, tableBody2_R);
                    }
                }
                startRow += 13;
            }

            aRow = outputSheet.getRow(startRow);
            aCell = aRow.getCell(3);
            aCell.value = '==========================================================================================';
            console.log('startRow', startRow);

            // 刪除style頁面
            this.workbook.removeWorksheet(worksheet.id);
            // 產生Excel檔案Buffer
            const buffer = await this.workbook.xlsx.writeBuffer();
            resolve(buffer);
        });
    }

    async getReportBuffer() {
        // 產生Excel檔案Buffer
        if (this.workbook != null) {
            const buffer = await this.workbook.xlsx.writeBuffer();
            // console.log(buffer);
            return buffer;
        } else {
            return null;
        }
    }

}

module.exports = ReportExcel;
