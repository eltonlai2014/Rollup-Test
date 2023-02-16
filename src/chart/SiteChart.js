const log4js = require('log4js');
const logger = log4js.getLogger('SiteChart');
const CommonChart = require('./CommonChart');
const _ = require('lodash');
// require('canvas-5-polyfill');
class SiteChart extends CommonChart {


    constructor(cWidth, cHeight, options) {
        super(cWidth, cHeight, options);
        this.myInit(options);
    }

    myInit(options) {
        logger.info('SiteChart init');
        // 方塊顏色
        this.chartColor = options.chartColor || ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#229B2F', '#6495ED'];
        this.chartLineWidth = options.chartLineWidth || 4;
    }

    setChartData(data, type) {
        // company data & site data
        type = type || SiteChart.Type.COMPANY;
        logger.info('SiteChart setChartData ...');
        super.setChartData(data);
        this.chartData = data.data;
        this.chartTitle = data.title;
        // let data = [
        //     {Month: '六月', Health: 97, Warning: 2, Critical: 1},
        //     {Month: '七月', Health: 97, Warning: 2, Critical: 1},
        //     {Month: '八月', Health: 97, Warning: 2, Critical: 1}
        // ]
        // 遍歷歷史資料計算最大值最小值
        this.maxValue = 0;
        for (let i = 0; i < this.chartData.length; i++) {
            let aInfo = this.chartData[i];
            if(type == SiteChart.Type.COMPANY) {
                aInfo.Total = aInfo.Health + aInfo.Warning + aInfo.Critical;
                this.maxValue = Math.max(this.maxValue, aInfo.Total);
            }            
            else if(type == SiteChart.Type.SITES) {
                this.maxValue = Math.max(this.maxValue, aInfo.Warning);
                this.maxValue = Math.max(this.maxValue, aInfo.Critical);
            }    
        }
        // 將資料依照日期排序
        this.chartData = _.sortBy(this.chartData, ['Month']);
        console.log('data', this.chartData, this.maxValue);
        // 取坐標軸最大值
        this.axisY_Max = this.getPrettyUnit(this.maxValue);
        return this;
    }

    getChartInfo() {
        console.log(this.leftWidth, this.chartWidth, this.rightWidth);
        console.log(this.topHeight, this.chartHeight, this.bottomHeight);
    }
    paint(key) {
        logger.info('SiteChart paint ...');
        key = key || ['Health', 'Warning', 'Critical', 'Total'];
        this.getChartInfo();
        const fontStyle_Normal = '';
        const fontStyle_Bold = 'bold';
        // 背景色
        const bgColor = '#ffffff';
        const aContext = this.context;
        aContext.fillStyle = bgColor;
        aContext.fillRect(0, 0, this.cWidth, this.cHeight);

        // 外框線
        const bolderColor = '#AAAAAA';
        const borderWidth = 1;
        this.clearLineTo(aContext, 0, 0, this.cWidth, 0, bolderColor, borderWidth);
        this.clearLineTo(aContext, 0, this.cHeight - 1, this.cWidth, this.cHeight - 1, bolderColor, borderWidth);
        this.clearLineTo(aContext, 0, 0, 0, this.cHeight - 1, bolderColor, borderWidth);
        this.clearLineTo(aContext, this.cWidth - 1, 0, this.cWidth - 1, this.cHeight - 1, bolderColor, borderWidth);

        // 圖標題
        if(this.chartTitle){
            const title_FontSize = 16 ;
            const title_Font = 'Arial' ;
            const title_Color = '#333333';
            this.drawString(aContext, this.chartTitle, this.cWidth / 2, this.topHeight / 2 + 4, title_FontSize, title_Font, fontStyle_Normal, title_Color, 'center', 'middle');    
        }

        // X Y軸線
        const axisColor = '#CCCCCC';
        const axisWidth = 1;
        this.clearLineTo(aContext, this.leftWidth - 1, this.topHeight, this.leftWidth - 1, this.topHeight + this.chartHeight, axisColor, axisWidth);
        this.clearLineTo(aContext, this.leftWidth - 1, this.topHeight + this.chartHeight, this.cWidth - this.rightWidth, this.topHeight + this.chartHeight, axisColor, axisWidth);

        // 畫Y軸座標與水平線
        const axisY_FontSize = 10;
        const yLines = 5;
        const label_Font = 'Arial';
        const label_Color = '#333333';
        for (let i = 0; i < yLines; i++) {
            const yPos = this.topHeight + i * this.chartHeight / yLines;
            if (i > 0) {
                // 水平線
                // this.dashedLineTo(aContext, this.leftWidth - 1, yPos, this.cWidth - this.rightWidth, yPos, axisColor, axisWidth);
                this.clearLineTo(aContext, this.leftWidth - 1, yPos, this.cWidth - this.rightWidth, yPos, axisColor, axisWidth);
            }
            // 座標
            this.drawString(aContext, (yLines - i) * this.axisY_Max / yLines, this.leftWidth - 4, yPos, axisY_FontSize, label_Font, fontStyle_Normal, label_Color, 'right', 'middle');
        }
        // 最小值座標
        this.drawString(aContext, '0', this.leftWidth - 4, this.topHeight + this.chartHeight + 2, axisY_FontSize, label_Font, fontStyle_Normal, label_Color, 'right', 'bottom');
        this.drawString(aContext, 'Sites', this.leftWidth - 4, this.topHeight + this.chartHeight + 18, axisY_FontSize, label_Font, fontStyle_Normal, label_Color, 'right', 'top');

        // 畫圖
        aContext.save();
        const unitWidth = this.chartWidth / this.chartData.length;
        // let key = ['Health', 'Warning', 'Critical', 'Total'];
        for (let j = 0; j < key.length; j++) {
            let aKey = key[j];
            if (this.chartData.length == 1) {
                // 只有一筆資料，畫長條圖
                let aInfo = this.chartData[0];
                const rectWidth = this.chartWidth / (key.length + 2);
                const xPos = this.leftWidth + (j + 1.5) * rectWidth;
                const rechHeight = (aInfo[aKey] / this.axisY_Max) * this.chartHeight;
                this.fillRectEx(aContext, xPos - rectWidth * 0.3, this.topHeight + this.chartHeight, rectWidth * 0.6, -rechHeight, this.chartColor[(j % this.chartColor.length)]);

                // 畫數值 Label
                const yPos = this.topHeight + this.chartHeight - rechHeight - 4;
                // this.drawString(aContext, aInfo[aKey], xPos, yPos, 10, label_Font, fontStyle_Normal, label_Color, 'center', 'bottom');
                this.drawBgString(aContext, aInfo[aKey], xPos, yPos, 10, label_Font, fontStyle_Normal, label_Color, bgColor, 'center', 'bottom');
                // 日期Label
                if (j == 0) {
                    this.drawString(aContext, aInfo.Month, this.leftWidth + this.chartWidth / 2, this.topHeight + this.chartHeight + 18, 10, label_Font, fontStyle_Normal, label_Color, 'center', 'top');
                }

            } else {
                // 多筆資料，繪製趨勢圖
                let fromX = null;
                let fromY = null;
                for (let i = 0; i < this.chartData.length; i++) {
                    let aInfo = this.chartData[i];
                    const xPos = this.leftWidth + (i + 0.5) * unitWidth;
                    let yPos = this.topHeight + ((this.axisY_Max - aInfo[aKey]) / this.axisY_Max) * this.chartHeight;
                    // console.log(xPos, yPos, aInfo.time);
                    if (fromX != null && fromY != null) {
                        this.clearLineTo(aContext, fromX, fromY, xPos, yPos, this.chartColor[(j % this.chartColor.length)], this.chartLineWidth);
                    }
                    fromX = xPos;
                    fromY = yPos;
                }
                let monthOffset = false;
                if (this.chartData.length > 6) {
                    monthOffset = true;
                }
                for (let i = 0; i < this.chartData.length; i++) {
                    let aInfo = this.chartData[i];
                    const xPos = this.leftWidth + (i + 0.5) * unitWidth;
                    let yPos = this.topHeight + ((this.axisY_Max - aInfo[aKey]) / this.axisY_Max) * this.chartHeight;
                    // 畫數值 Label
                    let offset = -4;
                    if (j % 2 == 0) {
                        offset = 20;
                    }
                    // this.drawString(aContext, aInfo[aKey], xPos, yPos , 10, label_Font, fontStyle_Normal, label_Color, 'center', 'middle');
                    this.drawBgString(aContext, aInfo[aKey], xPos, yPos + offset, 10, label_Font, fontStyle_Normal, this.chartColor[(j % this.chartColor.length)], bgColor, 'center', 'bottom');
                    // 日期Label
                    if (j == 0) {
                        let dateLabel = aInfo.Month.toString();
                        if(monthOffset) {
                            if(dateLabel.length == 6){
                                dateLabel = dateLabel.substring(4);
                            }
                        }
                        this.drawString(aContext, dateLabel, xPos, this.topHeight + this.chartHeight + 18, 10, label_Font, fontStyle_Normal, label_Color, 'center', 'top');
                    }
                }                
            }

        }

        // 畫圖例說明
        let hintWidth = 30;
        let hintHeight = 6;

        // 用圖例數量決定位置
        let yPos = this.cHeight - 16;
        for (let i = 0; i < key.length; i++) {
            let xPos = this.leftWidth + (this.chartWidth / key.length) * i;
            this.drawString(aContext, key[i], xPos + hintWidth + 4, yPos, 10, label_Font, fontStyle_Normal, label_Color, 'left', 'moddle');
            this.fillRectEx(aContext, xPos, yPos - hintHeight - 2, hintWidth, hintHeight, this.chartColor[i]);
        }

        aContext.restore();
        return this;
    }

}
// 常數定義
SiteChart.Type = {};
SiteChart.Type.COMPANY = 0;
SiteChart.Type.SITES = 1;
module.exports = SiteChart;
