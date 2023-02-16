const log4js = require('log4js');
const logger = log4js.getLogger('RectChart');
const CommonChart = require('./CommonChart');
const _ = require('lodash');
// require('canvas-5-polyfill');
class TrendChart extends CommonChart {

    constructor(cWidth, cHeight, options) {
        super(cWidth, cHeight, options);
        this.myInit(options);
    }

    myInit(options) {
        logger.info('TrendChart init');
        // 方塊顏色
        this.trendColor = options.rectColor || ['#808080', '#B0C4DE', '#E6E6FA', '#FFF0F5', '#229B2F', '#6495ED'];
    }

    setChartData(data) {
        logger.info('TrendChart setChartData ...');
        super.setChartData(data);
        this.chartData = data;
        // 遍歷歷史資料計算最大值最小值
        this.maxValue = 0;
        this.eventList = [];
        _.forEach(this.chartData, (value, key) => {
            // logger.info('key', key, 'value', value);
            this.eventList.push(key);
            this.maxValue = Math.max(this.maxValue, value.i);
            this.maxValue = Math.max(this.maxValue, value.o);
        });
        this.axisY_Max = this.getPrettyUnit(this.maxValue);
        this.startTime = 1661423000;
        this.endTime = 1661509860;
        this.timeUnit = 300;
        this.xUnit = this.chartWidth / (this.endTime - this.startTime);
        this.totalTime = this.endTime - this.startTime;
        return this;
    }

    setSiteId(site_id) {
        this.site_id = site_id;

        return this;
    }

    timeToX(aInfo) {
        let scale = ((aInfo.time - this.startTime) / (this.endTime - this.startTime));
        if (scale > 1) {
            console.log(aInfo.time, this.startTime, this.endTime);
        }
        return Math.round(((aInfo.time - this.startTime) / (this.endTime - this.startTime)) * this.chartWidth);
    }

    getChartInfo() {
        console.log(this.leftWidth, this.chartWidth, this.rightWidth);
        console.log(this.topHeight, this.chartHeight, this.bottomHeight);
    }
    paint() {
        logger.info('RectChart paint ...');
        this.getChartInfo();
        const fontStyle_Normal = '';
        const fontStyle_Bold = 'bold';
        // 背景色
        const aContext = this.context;
        aContext.fillStyle = '#ffffff';
        aContext.fillRect(0, 0, this.cWidth, this.cHeight);

        // 圖標題
        const title_FontSize = 16;
        const title_Font = 'Arial';
        const title_Color = '#000000';
        this.drawString(aContext, this.site_id, this.leftWidth, this.topHeight / 2, title_FontSize, title_Font, fontStyle_Normal, title_Color, 'left', 'middle');

        // X Y軸線
        const axisColor = '#000000';
        const axisWidth = 1;
        this.clearLineTo(aContext, this.leftWidth - 1, this.topHeight, this.leftWidth - 1, this.topHeight + this.chartHeight, axisWidth, axisColor, axisWidth);
        this.clearLineTo(aContext, this.leftWidth - 1, this.topHeight + this.chartHeight, this.cWidth - this.rightWidth, this.topHeight + this.chartHeight, axisColor, axisWidth);

        // 畫Y軸座標與水平線
        const axisY_FontSize = 10;
        const yLines = 5;
        const label_Font = 'Arial';
        const label_Color = '#000000';
        for (let i = 0; i < yLines; i++) {
            const yPos = this.topHeight + i * this.chartHeight / yLines;
            if (i > 0) {
                // 水平線
                this.dashedLineTo(aContext, this.leftWidth - 1, yPos, this.cWidth - this.rightWidth, yPos, axisColor, axisWidth);
            }
            // 座標
            this.drawString(aContext, (yLines - i) * this.axisY_Max / yLines, this.leftWidth - 4, yPos, axisY_FontSize, label_Font, fontStyle_Normal, label_Color, 'right', 'middle');
        }
        this.drawString(aContext, '0', this.leftWidth - 4, this.topHeight + this.chartHeight, axisY_FontSize, label_Font, fontStyle_Normal, label_Color, 'right', 'middle');

        // 畫趨勢圖
        aContext.save();
        let fromX = null;
        let fromY = null;
        for (let i = 0; i < this.chartData.length; i++) {
            let aInfo = this.chartData[i];
            let xPos = this.leftWidth + this.timeToX(aInfo);
            let yPos = this.topHeight + (aInfo.i / this.maxValue) * this.chartHeight;
            // console.log(xPos, yPos, aInfo.time);
            if (fromX != null && fromY != null) {
                this.clearLineTo(aContext, fromX, fromY, xPos, yPos, '#FF0000', 1);
            }
            fromX = xPos;
            fromY = yPos;
        }
        aContext.restore();

        return this;
    }

}
module.exports = TrendChart;
