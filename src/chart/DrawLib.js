class DrawLib {

    drawString(ctx, txt, x, y, size, font, fontStyle, color, align, base) {
        // 畫字串
        color = color || "#000000";
        base = base || "bottom";
        align = align || "left";
        font = font || "Arial";
        fontStyle = fontStyle || "";
        ctx.save();
        ctx.fillStyle = color;
        ctx.font = fontStyle + " " + size + "pt " + font;
        ctx.textAlign = align;
        ctx.textBaseline = base;
        x = Math.round(x);
        y = Math.round(y);
        ctx.fillText(txt, x, y);
        let wordWidth = ctx.measureText(txt).width;
        ctx.restore();
        return wordWidth;
    }


    drawBgString(ctx, txt, x, y, size, font, fontStyle, color, bgcolor, align, base) {
        // 繪製底色方塊 + 字串
        bgcolor = bgcolor || "#999999";
        fontStyle = fontStyle || "";
        ctx.font = fontStyle + " " + size + "pt " + font;
        // 計算字型寬/高
        const aWidth = ctx.measureText(txt).width;
        const aHeight = ctx.measureText('Ag').emHeightAscent;
        x = Math.round(x);
        y = Math.round(y);
        let rectPosY = y - aHeight - 3;
        let rectHeight = aHeight + 3;
        // 繪製底色方塊
        ctx.beginPath();
        if (align == "right") {
            ctx.rect(x - aWidth - 2, rectPosY, aWidth + 4, rectHeight);
        }
        else if (align == "left") {
            ctx.rect(x - 2, rectPosY, aWidth + 4, rectHeight);
        }
        else {
            ctx.rect(x - 2 - aWidth / 2, rectPosY, aWidth + 4, rectHeight);
        }
        ctx.fillStyle = bgcolor;
        ctx.fill();
        return this.drawString(ctx, txt, x, y, size, font, fontStyle, color, align, base) + 4;
    }

    drawBgStringRoundRect(ctx, txt, x, y, size, font, fontStyle, color, bgcolor, align, base) {
        // 繪製底色方塊 + 字串
        bgcolor = bgcolor || "#999999";
        fontStyle = fontStyle || "";
        ctx.font = fontStyle + " " + size + "pt " + font;
        // 計算字型寬/高
        const aWidth = ctx.measureText(txt).width;
        const aHeight = ctx.measureText('Ag').emHeightAscent;
        x = Math.round(x);
        y = Math.round(y);
        let rectPosY = y - aHeight - 3;
        let rectHeight = aHeight + 3;
        // 繪製底色方塊
        ctx.beginPath();
        if (align == "right") {
            this.drawRoundRect(ctx, x - aWidth - 2, rectPosY, aWidth + 4, rectHeight, 4, true, bgcolor);
        }
        else if (align == "left") {
            this.drawRoundRect(ctx, x - 2, rectPosY, aWidth + 4, rectHeight, 4, true, bgcolor);
        }
        else {
            this.drawRoundRect(ctx, x - 2 - aWidth / 2, rectPosY, aWidth + 4, rectHeight, 4, true, bgcolor);
        }
        ctx.fillStyle = bgcolor;
        ctx.fill();
        return this.drawString(ctx, txt, x, y, size, font, fontStyle, color, align, base) + 4;
    }

    // 基礎 method，暫無方法關閉 antialias，不建議用
    // drawLine(ctx, x1, y1, x2, y2, color, width) {
    //     width = width || 1;
    //     color = color || '#000000';
    //     ctx.beginPath();
    //     ctx.lineTo(x1, y1);
    //     ctx.lineTo(x2, y2);
    //     ctx.strokeStyle = color;
    //     ctx.stroke();
    // }

    // drawDashLine(ctx, x1, y1, x2, y2, color, width, dashStyle) {
    //     // 設置線條樣式
    //     dashStyle = dashStyle || [3, 3];
    //     ctx.setLineDash(dashStyle);
    //     this.drawLine(ctx, x1, y1, x2, y2, width, color);
    //     // 恢復實線
    //     ctx.setLineDash([]);
    // }

    // 圓角矩形
    drawRoundRect(ctx, x, y, width, height, radius, fill, fillStyle, stroke, lineWidth) {
        if (typeof stroke === "undefined") {
            stroke = false;
        }
        if (typeof fill === "undefined") {
            fill = false;
        }
        if (typeof radius === "undefined") {
            radius = 4;
        }
        lineWidth = lineWidth || 1;
        x = Math.round(x);
        y = Math.round(y);
        ctx.save();
        ctx.translate(0.5, 0.5);
        ctx.beginPath();
        // 畫圓角
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height - radius);
        ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        ctx.lineTo(x + radius, y + height);
        ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
        ctx.lineWidth = lineWidth;
        ctx.closePath();
        if (stroke) {
            ctx.stroke();
        }
        if (fill) {
            // let aColor = this.hexToRgb(mHint.BgColor[ColorSet]);
            // ctx.fillStyle = 'rgba(' + aColor.r + ',' + aColor.g + ',' + aColor.b + ',' + mHint.Alpha + ')';
            ctx.fillStyle = fillStyle;
            ctx.fill();
        }
        ctx.restore();
    }

    // 畫清晰虛線
    dashedLineTo(ctx, fromX, fromY, toX, toY, lineColor, lineWidth, pattern) {
        // 設置虛線樣式
        pattern = pattern || [3, 3];
        ctx.save();
        ctx.setLineDash(pattern);
        this.clearLineTo(ctx, fromX, fromY, toX, toY, lineColor, lineWidth);
        // 恢復實線
        ctx.restore();
    }

    // 畫清晰直線，避免antialias
    clearLineTo(ctx, fromX, fromY, toX, toY, lineColor, lineWidth) {
        lineWidth = lineWidth || 1;
        lineColor = lineColor || '#000000';
        // 避免畫線時產生antialias，save()->translate()->restore()
        ctx.save();
        ctx.translate(0.5, 0.5);
        // draw line
        fromX = Math.round(fromX);
        fromY = Math.round(fromY);
        toX = Math.round(toX);
        toY = Math.round(toY);
        ctx.beginPath();
        ctx.moveTo(fromX, fromY);
        ctx.lineTo(toX, toY);
        ctx.lineWidth = lineWidth;
        ctx.strokeStyle = lineColor;
        ctx.stroke();
        ctx.restore();
    }

    // 任意線段，避免antialias，用矩形模擬
    drawLineNoAliasing(ctx, sx, sy, tx, ty, lineColor) {
        lineColor = lineColor || '#FFFFFF';
        let dist = this.DBP(sx, sy, tx, ty);        // length of line
        let ang = this.getAngle(tx - sx, ty - sy);  // angle of line
        ctx.save();
        ctx.fillStyle = lineColor;
        for (let i = 0; i < dist; i++) {
            // for each point along the line
            ctx.fillRect(Math.round(sx + Math.cos(ang) * i), // round for perfect pixels
                Math.round(sy + Math.sin(ang) * i),          // thus no aliasing
                1, 1);                                       // fill in one pixel, 1x1
        }
        ctx.restore();
    }

    // 畫清晰矩形
    drawRectEx(ctx, x, y, width, height, color, lineWidth) {
        this.clearLineTo(ctx, x, y, x + width, y, color, lineWidth);
        this.clearLineTo(ctx, x + width, y, x + width, y + height, color, lineWidth);
        this.clearLineTo(ctx, x + width, y + height, x, y + height, color, lineWidth);
        this.clearLineTo(ctx, x, y + height, x, y, color, lineWidth);
    }

    fillRectEx(ctx, x, y, width, height, color) {
        ctx.beginPath();
        ctx.rect(Math.round(x), Math.round(y), Math.round(width), Math.round(height));
        ctx.fillStyle = color;
        ctx.fill();
        ctx.closePath();
    }

    // 座標取整數級距
    getPrettyUnit(value, aRatio) {
        // 最大值 放大比率
        let factor = 1.1;
        let unit = Math.pow(10, Math.floor(Math.log10(value)));
        let nextUnit = Math.pow(10, Math.ceil(Math.log10(value)));
        let ratio = value * factor / unit;
        // 小數字的處理
        if (nextUnit <= 10) {
            if (value < 4.5 * unit) {
                return nextUnit/2;
            }
            return nextUnit;
        }
        // 決定是否要換 下一個級距
        aRatio = aRatio || 7.5;
        if (ratio <= aRatio) {
            let ret = Math.ceil(value * factor / unit) * unit - unit / 2;
            return ret;
        }
        return nextUnit;
    }

    quickSort(arr) {
        if (arr.length <= 1) { return arr; }
        const pivotIndex = Math.floor(arr.length / 2);
        const pivot = arr.splice(pivotIndex, 1)[0];
        let left = [];
        let right = [];
        for (let i = 0, n = arr.length; i < n; i++) {
            if (arr[i] < pivot) {
                left.push(arr[i]);
            }
            else {
                right.push(arr[i]);
            }
        }
        return this.quickSort(left).concat([pivot], this.quickSort(right));
    }

    DBP(x1, y1, x2, y2) {
        return Math.sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1));
    }
    getAngle(x, y) {
        return Math.atan(y / (x === 0 ? 0.01 : x)) + (x < 0 ? Math.PI : 0);
    }
    /*
    d = 當前日期
    6 - w = 當前周的還有幾天過完（不算今天）
    兩者的和在除以7 就是當天是當前月份的第幾周
    */
    getMonthWeek(yyyy, MM, dd) {
        const date = new Date(yyyy, parseInt(MM, 10) - 1, dd);
        const w = date.getDay();
        const d = date.getDate();
        return Math.ceil((d + 6 - w) / 7);
    }
    /*
    date1是當前日期
    date2是當年第一天
    d是當前日期是今年第多少天
    用d + 當前年的第一天的周差距的和在除以7就是本年第幾周
    */
    getYearWeek(yyyy, MM, dd) {
        const date1 = new Date(yyyy, parseInt(MM) - 1, dd);
        const date2 = new Date(yyyy, 0, 1);
        const diff = Math.round((date1.valueOf() - date2.valueOf()) / 86400000);
        return Math.floor((diff + date2.getDay()) / 7);
    }
    rgbToHex(r, g, b) {
        return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
    }
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
        } : null;
    }

    setChartData(data) {
        console.log('CommonChart setChartData ...');
    }

}
module.exports = DrawLib;
