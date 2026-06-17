#target illustrator

function createGridAndCropMarks() {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてから実行してください。");
        return;
    }
    var doc = app.activeDocument;

    // --- ダイアログ表示処理 ---
    var dialog = new Window("dialog", "グリッド＆外トンボ生成");
    dialog.alignChildren = "fill";
    
    // 仕上がりサイズパネル
    var pSize = dialog.add("panel", undefined, "仕上がりサイズ (mm)");
    pSize.alignChildren = "right";
    pSize.margins = [15, 15, 15, 15];
    var gW = pSize.add("group"); gW.add("statictext", undefined, "幅:"); var inW = gW.add("edittext", undefined, "50"); inW.characters = 5;
    var gH = pSize.add("group"); gH.add("statictext", undefined, "高さ:"); var inH = gH.add("edittext", undefined, "50"); inH.characters = 5;

    // 配列設定パネル
    var pArray = dialog.add("panel", undefined, "配列設定");
    pArray.alignChildren = "right";
    pArray.margins = [15, 15, 15, 15];

    var gCols = pArray.add("group");
    gCols.add("statictext", undefined, "列数 (横):");
    var btnMinusCol = gCols.add("button", [0, 0, 25, 22], "-");
    var inCols = gCols.add("edittext", undefined, "2"); inCols.characters = 3;
    var btnPlusCol = gCols.add("button", [0, 0, 25, 22], "+");
    btnMinusCol.onClick = function() { var v = parseInt(inCols.text) || 1; if(v > 1) inCols.text = v - 1; };
    btnPlusCol.onClick = function() { var v = parseInt(inCols.text) || 1; inCols.text = v + 1; };

    var gRows = pArray.add("group");
    gRows.add("statictext", undefined, "行数 (縦):");
    var btnMinusRow = gRows.add("button", [0, 0, 25, 22], "-");
    var inRows = gRows.add("edittext", undefined, "2"); inRows.characters = 3;
    var btnPlusRow = gRows.add("button", [0, 0, 25, 22], "+");
    btnMinusRow.onClick = function() { var v = parseInt(inRows.text) || 1; if(v > 1) inRows.text = v - 1; };
    btnPlusRow.onClick = function() { var v = parseInt(inRows.text) || 1; inRows.text = v + 1; };

    var gHGap = pArray.add("group");
    gHGap.add("statictext", undefined, "横の間隔 (mm):");
    var btnMinusHGap = gHGap.add("button", [0, 0, 25, 22], "-");
    var inHGap = gHGap.add("edittext", undefined, "0"); inHGap.characters = 3;
    var btnPlusHGap = gHGap.add("button", [0, 0, 25, 22], "+");
    btnMinusHGap.onClick = function() { var v = parseFloat(inHGap.text) || 0; if(v > 0) inHGap.text = v - 1; };
    btnPlusHGap.onClick = function() { var v = parseFloat(inHGap.text) || 0; inHGap.text = v + 1; };

    var gVGap = pArray.add("group");
    gVGap.add("statictext", undefined, "縦の間隔 (mm):");
    var btnMinusVGap = gVGap.add("button", [0, 0, 25, 22], "-");
    var inVGap = gVGap.add("edittext", undefined, "0"); inVGap.characters = 3;
    var btnPlusVGap = gVGap.add("button", [0, 0, 25, 22], "+");
    btnMinusVGap.onClick = function() { var v = parseFloat(inVGap.text) || 0; if(v > 0) inVGap.text = v - 1; };
    btnPlusVGap.onClick = function() { var v = parseFloat(inVGap.text) || 0; inVGap.text = v + 1; };

    // トンボ設定パネル
    var pMark = dialog.add("panel", undefined, "トンボ設定");
    pMark.alignChildren = "right";
    pMark.margins = [15, 15, 15, 15];
    var gBleed = pMark.add("group"); gBleed.add("statictext", undefined, "塗り足し/マージン (mm):"); var inBleed = gBleed.add("edittext", undefined, "1"); inBleed.characters = 5;
    var gLen = pMark.add("group"); gLen.add("statictext", undefined, "トンボの長さ (mm):"); var inLen = gLen.add("edittext", undefined, "10"); inLen.characters = 5;
    var gStroke = pMark.add("group"); gStroke.add("statictext", undefined, "線の太さ (pt):"); var inStroke = gStroke.add("edittext", undefined, "0.5"); inStroke.characters = 5;

    // ボタン
    var btnGroup = dialog.add("group");
    btnGroup.alignment = "center";
    btnGroup.add("button", undefined, "OK", {name: "ok"});
    btnGroup.add("button", undefined, "キャンセル", {name: "cancel"});

    if (dialog.show() !== 1) return;

    // 入力値の取得と単位変換
    var mm = 2.834645;
    var wPt = parseFloat(inW.text) * mm;
    var hPt = parseFloat(inH.text) * mm;
    var cols = parseInt(inCols.text);
    var rows = parseInt(inRows.text);
    var hGapPt = parseFloat(inHGap.text) * mm;
    var vGapPt = parseFloat(inVGap.text) * mm;
    var bleedPt = parseFloat(inBleed.text) * mm;
    var markLenPt = parseFloat(inLen.text) * mm;
    var strokeW = parseFloat(inStroke.text);

    // 不正値の簡易チェック
    if (isNaN(wPt) || isNaN(hPt) || isNaN(cols) || isNaN(rows)) {
        alert("数値が正しく入力されていません。");
        return;
    }

    var totalWidth = (cols * wPt) + ((cols - 1) * hGapPt);
    var totalHeight = (rows * hPt) + ((rows - 1) * vGapPt);

    // 画面の中央に生成するための基準座標（左上）を計算
    var viewCenter = doc.views[0].centerPoint; // [x, y]
    var startX = viewCenter[0] - (totalWidth / 2);
    var startY = viewCenter[1] + (totalHeight / 2); // IllustratorのY軸は上がプラス

    var blackColor = new CMYKColor();
    blackColor.cyan = 0; blackColor.magenta = 0; blackColor.yellow = 0; blackColor.black = 100;

    var mainGroup = doc.groupItems.add();
    mainGroup.name = "Grid and CropMarks";

    // 1. 透明な矩形（仕上がりサイズ）をグリッド状に描画
    var rectGroup = mainGroup.groupItems.add();
    rectGroup.name = "CutLines";
    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var x = startX + c * (wPt + hGapPt);
            var y = startY - r * (hPt + vGapPt);
            
            // rectangle(top, left, width, height)
            var rect = rectGroup.pathItems.rectangle(y, x, wPt, hPt);
            rect.filled = false;
            rect.stroked = false; // 透明（線なし・塗りなし）に設定
        }
    }

    // 2. トンボの描画
    var markGroup = mainGroup.groupItems.add();
    markGroup.name = "CropMarks";

    // 重複する線を防ぐための仕組み
    var uniqueLines = {};
    function addUniqueLine(x1, y1, x2, y2) {
        // 微小な誤差を無視するために丸める
        var rx1 = Math.round(x1 * 1000) / 1000;
        var ry1 = Math.round(y1 * 1000) / 1000;
        var rx2 = Math.round(x2 * 1000) / 1000;
        var ry2 = Math.round(y2 * 1000) / 1000;
        
        var pt1 = rx1 + "," + ry1;
        var pt2 = rx2 + "," + ry2;
        var key = (pt1 < pt2) ? (pt1 + "_" + pt2) : (pt2 + "_" + pt1);
        
        if (!uniqueLines[key]) {
            uniqueLines[key] = [[x1, y1], [x2, y2]];
        }
    }

    // 行ごとのトンボ（左右辺）
    for (var r = 0; r < rows; r++) {
        var rectTop = startY - r * (hPt + vGapPt);
        var rectBottom = rectTop - hPt;

        // 上辺のライン（左外側と右外側）
        addUniqueLine(startX - bleedPt, rectTop, startX - bleedPt - markLenPt, rectTop);
        addUniqueLine(startX + totalWidth + bleedPt, rectTop, startX + totalWidth + bleedPt + markLenPt, rectTop);
        
        // 下辺のライン（左外側と右外側）
        addUniqueLine(startX - bleedPt, rectBottom, startX - bleedPt - markLenPt, rectBottom);
        addUniqueLine(startX + totalWidth + bleedPt, rectBottom, startX + totalWidth + bleedPt + markLenPt, rectBottom);
    }

    // 列ごとのトンボ（上下辺）
    for (var c = 0; c < cols; c++) {
        var rectLeft = startX + c * (wPt + hGapPt);
        var rectRight = rectLeft + wPt;

        // 左辺のライン（上外側と下外側）
        addUniqueLine(rectLeft, startY + bleedPt, rectLeft, startY + bleedPt + markLenPt);
        addUniqueLine(rectLeft, startY - totalHeight - bleedPt, rectLeft, startY - totalHeight - bleedPt - markLenPt);

        // 右辺のライン（上外側と下外側）
        addUniqueLine(rectRight, startY + bleedPt, rectRight, startY + bleedPt + markLenPt);
        addUniqueLine(rectRight, startY - totalHeight - bleedPt, rectRight, startY - totalHeight - bleedPt - markLenPt);
    }

    // トンボの線を黒色・指定ptに設定
    for (var key in uniqueLines) {
        var p = markGroup.pathItems.add();
        p.setEntirePath(uniqueLines[key]);
        p.filled = false;
        p.stroked = true;
        p.strokeColor = blackColor;
        p.strokeWidth = strokeW;
    }
}

createGridAndCropMarks();