const log4js = require('log4js');
const logger = log4js.getLogger('CommonChart');
const { createCanvas, loadImage } = require('canvas');
const DrawLib = require('./DrawLib');
class CommonChart extends DrawLib {

    constructor(cWidth, cHeight, options) {
        super();
        this.cWidth = cWidth;
        this.cHeight = cHeight;
        this.init(options);
    }

    init(options) {
        logger.info('CommonChart init');
        // 計算畫面大小
        this.leftWidth = options.leftWidth || 50;
        this.rightWidth = options.rightWidth || 20;
        this.chartWidth = this.cWidth - this.leftWidth - this.rightWidth;

        this.topHeight = options.topHeight || 50;
        this.bottomHeight = options.bottomHeight || 20;
        this.chartHeight = this.cHeight - this.topHeight - this.bottomHeight;

        this.canvas = createCanvas(this.cWidth, this.cHeight);
        this.context = this.canvas.getContext('2d');
    }

    getChartInfo() {
        return this.chartWidth + " " + this.chartHeight
    }
    // 取得畫布 buffer資料
    getCanvasBuffer(imageType) {
        imageType = imageType || 'image/png';
        return this.canvas.toBuffer(imageType);
    }

    setChartData(data) {
        console.log('CommonChart setChartData ...');
    }

}
module.exports = CommonChart;
