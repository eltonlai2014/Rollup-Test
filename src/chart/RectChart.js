const log4js = require('log4js');
const logger = log4js.getLogger('RectChart');
const CommonChart = require('./CommonChart');
const _ = require('lodash');
// require('canvas-5-polyfill');
class RectChart extends CommonChart {

    constructor(cWidth, cHeight, options) {
        super(cWidth, cHeight, options);
        this.myInit(options);
    }

    myInit(options) {
        logger.info('RectChart init');
        // 方塊顏色
        this.rectColor = options.rectColor || ['#808080', '#B0C4DE', '#E6E6FA', '#FFF0F5', '#229B2F', '#6495ED'];
    }

    setChartData(data) {
        logger.info('RectChart setChartData ...');
        super.setChartData(data);
        this.chartData = data;
        // 遍歷各 event 計算最大值
        this.maxEventCount = 0;
        this.eventList = [];
        _.forEach(this.chartData, (value, key) => {
            // logger.info('key', key, 'value', value);
            this.eventList.push(key);
            this.maxEventCount = Math.max(this.maxEventCount, value.count);
        });
        this.axisY_Max = this.getPrettyUnit(this.maxEventCount);
        return this;
    }

    setSiteId(site_id) {
        this.site_id = site_id;
        return this;
    }

    paint() {
        logger.info('RectChart paint ...');
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

        // 畫方塊
        const unitWidth = this.chartWidth / (this.eventList.length + 1);
        const rectWidth = unitWidth * 0.6;
        const eventLabelYPos = this.topHeight + this.chartHeight + this.bottomHeight / 2;
        for (let i = 0; i < this.eventList.length; i++) {
            const aInfo = this.chartData[this.eventList[i]];
            const rechHeight = (aInfo.count / this.axisY_Max) * this.chartHeight;
            const xPos = this.leftWidth + (i + 1) * unitWidth;
            this.fillRectEx(aContext, xPos - rectWidth / 2, this.topHeight + this.chartHeight, rectWidth, -rechHeight, this.rectColor[(i % this.rectColor.length)]);
            // this.clearLineTo(aContext, xPos, this.topHeight, xPos, this.topHeight + this.chartHeight);
            // 數值 Label
            // drawBgString(ctx, txt, x, y, size, font, color, bgcolor, align, base)
            if (aInfo.count == this.maxEventCount && this.maxEventCount > 0) {
                this.drawBgString(aContext, aInfo.count, xPos, this.topHeight + this.chartHeight - rechHeight - 2, 10, label_Font, fontStyle_Bold, '#FFFFFF', '#CC0000', 'center', 'bottom');
            } else {
                this.drawBgString(aContext, aInfo.count, xPos, this.topHeight + this.chartHeight - rechHeight - 2, 10, label_Font, fontStyle_Normal, label_Color, '#FFFFFF', 'center', 'bottom');
            }

            // 畫 eventType Label
            this.drawString(aContext, this.eventList[i], xPos, eventLabelYPos, 10, label_Font, fontStyle_Normal, label_Color, 'center', 'middle');
        }
        return this;
    }

}
module.exports = RectChart;
