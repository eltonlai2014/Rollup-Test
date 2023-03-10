const _ = require('lodash');
const log4js = require('log4js');
const logger = log4js.getLogger('ReportExcel');

const Excel = require('exceljs');
const SiteChect = require('./chart/SiteChart');

class ReportExcelSite {

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
            // ????????????????????????????????????
            let totalRows = 11;                                                         // 11 title block
            let OperationBlock = this.getParamValue('OperationBlock', true);            // 19
            let NetworkDeviceBlock = this.getParamValue('NetworkDeviceBlock', true);    // 19
            let EventTypeBlock = this.getParamValue('EventTypeBlock', true);            // 17
            let EventRankBlock = this.getParamValue('EventRankBlock', true);            // 17
            let WarningEventBlock = this.getParamValue('WarningEventBlock', true);      // 41
            if(OperationBlock) {
                // ????????????????????????????????????
                totalRows+= (16 + report.reportData.data.length);
            }
            if(NetworkDeviceBlock) {
                // ????????????????????????????????????
                totalRows+= (16 + report.reportData2.data.length);
            }
            if(EventTypeBlock || EventRankBlock) {
                totalRows+= 17;
            }
            if(WarningEventBlock) {
                totalRows+= 41;
            }

            // ??????????????????
            this.workbook = new Excel.Workbook();
            await this.workbook.xlsx.readFile(this.options.StyleFile);
            const worksheet = this.workbook.getWorksheet(this.options.StyleSheet);

            // ??????????????????
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
            // [??????????????????]
            const aSiteChect = new SiteChect(this.options.chartOptions.chartWidth || 540, this.options.chartOptions.chartHeight || 300, this.options.chartOptions);

            // ??????????????????
            let colWidth = [];
            for (let i = 1; i <= 20; i++) {
                let aColumn = worksheet.getColumn(i);
                colWidth.push(aColumn.width);
            }
            console.log('colWidth', colWidth);
            // ????????????
            // ????????????????????????????????????????????????

            // ????????????????????????
            const bgGray = worksheet.getCell('B1').style;
            const bgWhite = worksheet.getCell('B2').style;

            const reportTitle = worksheet.getCell('H3').style;
            const reportSummary = worksheet.getCell('C7').style;
            const blockTitle = worksheet.getCell('C13').style;
            const blockContent = worksheet.getCell('F13').style;
            const tableTitle_L = worksheet.getCell('C23').style;
            const tableBody_L = worksheet.getCell('C24').style;
            const tableTitle_R = worksheet.getCell('D23').style;
            const tableBody_R = worksheet.getCell('D24').style;

            const tableTitle2_L = worksheet.getCell('C53').style;
            const tableTitle2_R = worksheet.getCell('E53').style;
            const tableFooter2_L = worksheet.getCell('C63').style;
            const tableFooter2_R = worksheet.getCell('E63').style;            
            const tableBody2_L = worksheet.getCell('C54').style;
            const tableBody2_R = worksheet.getCell('E54').style;
            const tableBody2_L_even = worksheet.getCell('C55').style;
            const tableBody2_R_even = worksheet.getCell('E55').style;

            const tableTitle3_L = worksheet.getCell('G53').style;
            const tableTitle3_R = worksheet.getCell('M53').style;
            const tableFooter3_L = worksheet.getCell('G64').style;
            const tableFooter3_R = worksheet.getCell('M64').style;            
            const tableBody3_L = worksheet.getCell('G54').style;
            const tableBody3_R = worksheet.getCell('M54').style;
            const tableBody3_L_even = worksheet.getCell('G55').style;
            const tableBody3_R_even = worksheet.getCell('M55').style;

            const ServiceSummaryRows = 6;

            // ?????????????????????
            let outputSheet = this.workbook.addWorksheet(report.siteName, {
                pageSetup: { paperSize: 12 }
            });

            // ??????????????????
            for (let i = 1; i <= 20; i++) {
                let aColumn = outputSheet.getColumn(i);
                aColumn.width = colWidth[i - 1];
            }

            const BG_ROW_COUNT = 200;
            const RP_ROW_COUNT = totalRows + 1;
            // RP_ROW_COUNT ???????????????????????????????????????

            // ?????????????????????
            for (let i = 1; i <= BG_ROW_COUNT; i++) {
                let aRow = outputSheet.getRow(i);
                // ??????
                for (let j = 1; j <= 20; j++) {
                    let aCell = aRow.getCell(j);
                    aCell.style = bgGray;
                }
                // ??????
                if (i > 1 && i < RP_ROW_COUNT) {
                    for (let j = 2; j <= 14; j++) {
                        let aCell = aRow.getCell(j);
                        aCell.style = bgWhite;
                    }
                }
            }

            // [????????????]
            let startRow = 3;
            let aRow = outputSheet.getRow(startRow);
            let aCell = aRow.getCell(8);
            aCell.value = `???????????? _ ${report.siteName}`;
            aCell.style = reportTitle;
            aCell.font = reportTitle.font;
            startRow += 2;

            // ====================================================================================================
            // ??????Summary
            // ====================================================================================================
            const SummaryRows = 6;
            for (let i = startRow; i < startRow + SummaryRows; i++) {
                let aRow = outputSheet.getRow(i);
                // ??????????????????
                for (let i = 3; i <= 13; i++) {
                    let aCell = aRow.getCell(i);
                    aCell.style = reportSummary;
                }
                switch (i) {
                    case startRow + 1: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  ?????????${report.companyName}`;
                        break;
                    }
                    case startRow + 2: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  ?????????${report.siteName}`;
                        break;
                    }                    
                    case startRow + 3: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  ?????????????????????${report.reportDate}`;
                        break;
                    }
                    case startRow + 4: {
                        outputSheet.mergeCells(`C${i}:F${i}`);
                        let aCell = aRow.getCell(3);
                        aCell.value = `  ?????????????????????${report.reportSDate}~${report.reportEate}`;
                        break;
                    }
                }
            }
            startRow += SummaryRows;
            startRow += 2;

            // ====================================================================================================
            // ????????????????????????
            // ====================================================================================================
            if(OperationBlock) {
                aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:M${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = '????????????????????????';
                aCell.style = blockTitle;
                const imgStartRow = startRow + 1;
                startRow += 2;
                // [??????????????????]
                const expressLabel = [`  ??????????????? MXview One Server Alert ??????`, `  Critical Site  ??????????????? Critical Event`,
                    `  Warning Site  ???????????????Critical Event ?????? Warning Event`, `  Health Site  ??????????????? Critical Event ??? Warning Event`,
                ];
                for (let i = startRow; i < startRow + ServiceSummaryRows; i++) {
                    let aRow = outputSheet.getRow(i);
                    // ??????????????????
                    for (let j = 3; j <= 7; j++) {
                        let aCell = aRow.getCell(j);
                        aCell.style = reportSummary;
                    }

                    // ???????????????
                    switch (i) {
                        case startRow + 1:
                        case startRow + 2:
                        case startRow + 3:
                        case startRow + 4:
                            outputSheet.mergeCells(`C${i}:G${i}`);
                            let aCell = aRow.getCell(3);
                            aCell.value = expressLabel[i - startRow - 1];
                            break;
                    }
                }
                startRow += ServiceSummaryRows;
                startRow += 2;

                // [????????????]
                let title = ['', 'Warning Event', 'Critical Event', '','Site Status'];
                for (let i = 0; i < report.reportData.data.length; i++) {
                    let aInfo = report.reportData.data[i];
                    aInfo.Total = aInfo.Warning + aInfo.Critical;
                }
                // ???????????????????????????
                report.reportData.data = _.sortBy(report.reportData.data, ['Month']);

                for (let i = startRow; i <= startRow + (report.reportData.data.length + 1); i++) {
                    let aRow = outputSheet.getRow(i);
                    outputSheet.mergeCells(`F${i}:G${i}`);
                    switch (i) {
                        case startRow: {
                            // ????????????
                            for (let j = 0; j < 5; j++) {
                                let aCell = aRow.getCell(j + 3);
                                aCell.value = title[j];
                                aCell.style = tableTitle_R;
                            }
                            break;
                        }
                        default:
                            // ??????
                            let aInfo = report.reportData.data[i - startRow - 1];
                            if (aInfo) {
                                let aCell = aRow.getCell(3);
                                aCell.style = tableBody_L;
                                aCell.value = aInfo.Month;
                                aCell = aRow.getCell(4);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Warning;
                                aCell = aRow.getCell(5);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Critical;
                                aCell = aRow.getCell(6);
                                aCell.style = tableBody_R;
                                let siteStatus = 'Health Site';                             
                                if(aInfo.Warning > 0) {
                                    siteStatus = 'Warning Site'
                                }
                                if(aInfo.Critical > 0) {
                                    siteStatus = 'Critical Site'
                                }                                   
                                aCell.value = siteStatus;
                            }

                            break;
                    }
                }

                // [??????]
                // ?????????????????????png buffer
                aSiteChect.setChartData(report.reportData, SiteChect.Type.SITES).paint(['Warning', 'Critical']);
                const buffer = aSiteChect.getCanvasBuffer();

                // ??????excel workbook
                const imageId1 = this.workbook.addImage({
                    buffer: buffer,
                    extension: 'png',
                });
                outputSheet.addImage(imageId1, { tl: { col: 7.5, row: imgStartRow }, ext: { width: this.options.chartOptions.chartWidth, height: this.options.chartOptions.chartHeight } });

                startRow += report.reportData.data.length;
                startRow += 5;
            }

            // ====================================================================================================
            // ?????????????????????????????????
            // ====================================================================================================
            if(NetworkDeviceBlock) {
                aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:M${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = '?????????????????????????????????';
                aCell.style = blockTitle;
                const imgStartRow = startRow + 1;
                startRow += 2;

                // [??????????????????]
                const SiteSummaryRows = 6;
                const expressLabel_2 = [`  ??????????????? MXview??One Server Alert ????????? Alert ??????`, `  Critical Site  ??????????????? Critical Event`,
                    `  Warning Site  ???????????????Critical Event ?????? Warning Event`, `  Health Site  ??????????????? Critical Event ??? Warning Event`,
                ];
                for (let i = startRow; i < startRow + SiteSummaryRows; i++) {
                    let aRow = outputSheet.getRow(i);
                    // ??????????????????
                    for (let j = 3; j <= 7; j++) {
                        let aCell = aRow.getCell(j);
                        aCell.style = reportSummary;
                    }

                    // ???????????????
                    switch (i) {
                        case startRow + 1:
                        case startRow + 2:
                        case startRow + 3:
                        case startRow + 4:
                            outputSheet.mergeCells(`C${i}:G${i}`);
                            let aCell = aRow.getCell(3);
                            aCell.value = expressLabel_2[i - startRow - 1];
                            break;
                    }
                }
                startRow += ServiceSummaryRows;
                startRow += 2;

                // [????????????]
                let title_2 = ['', 'Warning Site', 'Critical Site', '', 'Site Status'];
                for (let i = 0; i < report.reportData2.data.length; i++) {
                    let aInfo = report.reportData2.data[i];
                    aInfo.Total = aInfo.Health + aInfo.Warning + aInfo.Critical;
                }
                // ???????????????????????????
                report.reportData2.data = _.sortBy(report.reportData2.data, ['Month']);

                for (let i = startRow; i <= startRow + (report.reportData2.data.length + 1); i++) {
                    let aRow = outputSheet.getRow(i);
                    outputSheet.mergeCells(`F${i}:G${i}`);
                    switch (i) {
                        case startRow: {
                            // ????????????
                            for (let j = 0; j < 5; j++) {
                                let aCell = aRow.getCell(j + 3);
                                aCell.value = title_2[j];
                                aCell.style = tableTitle_R;
                            }
                            break;
                        }
                        default:
                            // ??????
                            let aInfo = report.reportData2.data[i - startRow - 1];
                            if (aInfo) {
                                let aCell = aRow.getCell(3);
                                aCell.style = tableBody_L;
                                aCell.value = aInfo.Month;
                                aCell = aRow.getCell(4);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Warning;
                                aCell = aRow.getCell(5);
                                aCell.style = tableBody_R;
                                aCell.value = aInfo.Critical;
                                aCell = aRow.getCell(6);
                                aCell.style = tableBody_R;
                                let siteStatus = 'Health Site';                             
                                if(aInfo.Warning > 0) {
                                    siteStatus = 'Warning Site'
                                }
                                if(aInfo.Critical > 0) {
                                    siteStatus = 'Critical Site'
                                }                                   
                                aCell.value = siteStatus;
                            }

                            break;
                    }
                }

                // [??????]
                // ?????????????????????png buffer
                aSiteChect.setChartData(report.reportData2, SiteChect.Type.SITES).paint(['Warning', 'Critical']);
                const buffer2 = aSiteChect.getCanvasBuffer();

                // ??????excel workbook
                const imageId2 = this.workbook.addImage({
                    buffer: buffer2,
                    extension: 'png',
                });
                outputSheet.addImage(imageId2, { tl: { col: 7.5, row: imgStartRow }, ext: { width: this.options.chartOptions.chartWidth, height: this.options.chartOptions.chartHeight } });
                startRow += report.reportData2.data.length;
                startRow += 6;
            }

            let blockStartRow = startRow;
            // ====================================================================================================
            //  ????????????????????????
            // ====================================================================================================
            if(EventTypeBlock) {
                let aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`C${startRow}:E${startRow}`);
                aCell = aRow.getCell(3);
                aCell.value = ' ????????????????????????';
                aCell.style = blockTitle;
                startRow += 2;

                // [EventType 1~10]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 3, 'EventType', tableTitle2_L);
                        fillTitleCell(aRow, 4, '', tableTitle2_L);
                        fillTitleCell(aRow, 5, 'Count', tableTitle2_R);
                    } else {
                        let rankInfo = report.reportData3.RankList[i - startRow - 1];
                        fillCell(aRow, i - startRow, 3, rankInfo.EventName, tableBody2_L_even, tableBody2_L);
                        fillCell(aRow, i - startRow, 4, rankInfo.EventName, tableBody2_R_even, tableBody2_R);
                        fillCell(aRow, i - startRow, 5, rankInfo.Count, tableBody2_R_even, tableBody2_R);
                    }
                    outputSheet.mergeCells(`C${i}:D${i}`);
                }
                startRow += 11;
                aRow = outputSheet.getRow(startRow);
                fillTitleCell(aRow, 3, 'Total', tableFooter2_L);
                fillTitleCell(aRow, 4, '', tableFooter2_R);
                fillTitleCell(aRow, 5, report.reportData3.RankTotal, tableFooter2_R);
                outputSheet.mergeCells(`C${startRow}:D${startRow}`);
                startRow += 3;

            }

            // ====================================================================================================
            //  ?????????????????? Top 10
            // ====================================================================================================
            if(EventRankBlock) {
                // ????????????????????????
                startRow = blockStartRow;                
                let aRow = outputSheet.getRow(startRow);
                outputSheet.mergeCells(`G${startRow}:M${startRow}`);
                aCell = aRow.getCell(7);
                aCell.value = ' ?????????????????? Top 10';
                aCell.style = blockTitle;
                startRow += 2;

                // [EventType 1~10]
                for (let i = startRow; i <= startRow + 10; i++) {
                    let aRow = outputSheet.getRow(i);
                    if (i == startRow) {
                        fillTitleCell(aRow, 7, 'Categroy', tableTitle3_L);
                        fillTitleCell(aRow, 9, 'EventType', tableTitle3_L);
                        fillTitleCell(aRow, 13, 'Count', tableTitle3_R);
                    } else {
                        let rankInfo = report.reportData4.RankList[i - startRow - 1];
                        fillCell(aRow, i - startRow, 7, rankInfo.Categroy, tableBody3_L_even, tableBody3_L);
                        fillCell(aRow, i - startRow, 9, rankInfo.EventName, tableBody3_L_even, tableBody3_L);
                        fillCell(aRow, i - startRow, 13, rankInfo.Count, tableBody3_R_even, tableBody3_R);
                    }
                    outputSheet.mergeCells(`G${i}:H${i}`);
                    outputSheet.mergeCells(`I${i}:L${i}`);
                }
                startRow += 11;
                aRow = outputSheet.getRow(startRow);
                fillTitleCell(aRow, 7, 'Total', tableFooter3_L);
                fillTitleCell(aRow, 9, '', tableFooter3_R);
                fillTitleCell(aRow, 13, report.reportData4.RankTotal, tableFooter3_R);
                outputSheet.mergeCells(`G${startRow}:H${startRow}`);
                outputSheet.mergeCells(`I${startRow}:L${startRow}`);
                startRow += 3;
            }


            // ====================================================================================================
            // Warning Event ???????????? Top 10
            // ====================================================================================================
            if(WarningEventBlock) {

            }

            aRow = outputSheet.getRow(startRow);
            aCell = aRow.getCell(3);
            aCell.value = '==========================================================================================';
            console.log('startRow', startRow);

            // ??????style??????
            this.workbook.removeWorksheet(worksheet.id);
            // ??????Excel??????Buffer
            const buffer = await this.workbook.xlsx.writeBuffer();
            resolve(buffer);
        });
    }

    async getReportBuffer() {
        // ??????Excel??????Buffer
        if (this.workbook != null) {
            const buffer = await this.workbook.xlsx.writeBuffer();
            // console.log(buffer);
            return buffer;
        } else {
            return null;
        }
    }
}

module.exports = ReportExcelSite;
