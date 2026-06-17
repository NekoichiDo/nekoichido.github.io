#target illustrator

function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    var sel = doc.selection;

    // オブジェクトが2つ以上選択されていない場合は終了
    if (sel.length < 2) {
        alert("基準となる枠と、並べ替えるオブジェクトを選択してください。");
        return;
    }

    // --- ダイアログ表示処理 ---
    var dialog = new Window("dialog", "配置設定");
    dialog.alignChildren = "fill";
    
    var panel = dialog.add("panel", undefined, "間隔とオプション");
    panel.alignChildren = "right";
    
    var groupX = panel.add("group");
    groupX.add("statictext", undefined, "横の間隔 (mm):");
    var inputX = groupX.add("edittext", undefined, "3");
    inputX.characters = 5;

    var groupY = panel.add("group");
    groupY.add("statictext", undefined, "縦の間隔 (mm):");
    var inputY = groupY.add("edittext", undefined, "3");
    inputY.characters = 5;

    var groupCheck = panel.add("group");
    groupCheck.alignment = "right";
    var checkRemoveFrame = groupCheck.add("checkbox", undefined, "整列後に基準枠を消去する");
    checkRemoveFrame.value = true; // デフォルトは「消去する」

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, "OK", {name: "ok"});
    buttonGroup.add("button", undefined, "キャンセル", {name: "cancel"});

    // キャンセルされた場合は終了
    if (dialog.show() !== 1) {
        return;
    }

    // 入力値の取得と数値化
    var gapX_mm = parseFloat(inputX.text);
    var gapY_mm = parseFloat(inputY.text);
    var removeFrame = checkRemoveFrame.value;

    // 不正な値が入力された場合は0にする
    if (isNaN(gapX_mm)) gapX_mm = 0;
    if (isNaN(gapY_mm)) gapY_mm = 0;
    // ------------------------

    // ▼▼▼ マスクのサイズ（見た目のサイズ）を正確に取得する処理 ▼▼▼
    function getTrueBounds(obj) {
        if (obj.typename === "GroupItem") {
            if (obj.clipped) {
                // クリッピンググループ内の先頭アイテムをマスクとして扱う
                if (obj.pageItems.length > 0) {
                    return obj.pageItems[0].geometricBounds;
                }
            } else {
                // 通常のグループの場合、中身をすべてチェックしてサイズを合成する
                var L = Infinity, T = -Infinity, R = -Infinity, B = Infinity;
                var hasValidBounds = false;
                for (var i = 0; i < obj.pageItems.length; i++) {
                    var b = getTrueBounds(obj.pageItems[i]); // 再帰処理
                    if (b) {
                        L = Math.min(L, b[0]);
                        T = Math.max(T, b[1]);
                        R = Math.max(R, b[2]);
                        B = Math.min(B, b[3]);
                        hasValidBounds = true;
                    }
                }
                if (hasValidBounds) return [L, T, R, B];
            }
        }
        
        // グループ以外の場合
        try {
            return obj.geometricBounds;
        } catch(e) {
            return null;
        }
    }

    function getBoundsInfo(obj) {
        var b = getTrueBounds(obj);
        if (!b) {
            b = obj.geometricBounds; // 万が一のための保険
        }
        // boundsは [left, top, right, bottom]
        return {
            obj: obj,
            left: b[0],
            top: b[1],
            right: b[2],
            bottom: b[3],
            width: b[2] - b[0],
            height: b[1] - b[3]
        };
    }
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // 全選択オブジェクトのサイズ情報を取得
    var itemsInfo = [];
    for (var i = 0; i < sel.length; i++) {
        itemsInfo.push(getBoundsInfo(sel[i]));
    }

    // 選択範囲の中で最も幅が広いオブジェクトを「基準の枠」と判定
    var frameInfo = itemsInfo[0];
    for (var i = 1; i < itemsInfo.length; i++) {
        if (itemsInfo[i].width > frameInfo.width) {
            frameInfo = itemsInfo[i];
        }
    }

    // 基準枠以外のオブジェクトを配列にまとめる
    var targets = [];
    for (var i = 0; i < itemsInfo.length; i++) {
        if (itemsInfo[i].obj !== frameInfo.obj) {
            targets.push(itemsInfo[i]);
        }
    }

    // 左上にあるものから順にソート
    // Y座標の差が5pt以内なら同じ行と見なしてX座標でソート
    targets.sort(function(a, b) {
        var yDiff = b.top - a.top; 
        if (Math.abs(yDiff) > 5) {
            return yDiff; // 上にあるほうが先
        }
        return a.left - b.left; // 左にあるほうが先
    });

    // 配置処理の初期設定
    var startX = frameInfo.left;
    var startY = frameInfo.top;
    var maxWidth = frameInfo.width;
    
    // 間隔（ミリメートルをIllustratorの内部単位 pt に変換）
    var PT_PER_MM = 2.834645;
    var gapX = gapX_mm * PT_PER_MM;
    var gapY = gapY_mm * PT_PER_MM;

    var currentX = startX;
    var currentY = startY;
    var rowHeight = 0;

    // 順番に配置していく
    for (var i = 0; i < targets.length; i++) {
        var info = targets[i];

        // 行の幅を超える場合は改行（行の最初の要素ははみ出しても配置する）
        if (currentX !== startX && (currentX - startX + info.width > maxWidth)) {
            currentX = startX;
            currentY = currentY - rowHeight - gapY;
            rowHeight = 0;
        }

        // オブジェクトを移動（マスクのはみ出しを考慮して、目標座標との「差分」で移動させる）
        var deltaX = currentX - info.left;
        var deltaY = currentY - info.top;
        info.obj.translate(deltaX, deltaY);

        // 次のオブジェクトのX座標と、現在の行の最大の高さを更新
        currentX += info.width + gapX;
        rowHeight = Math.max(rowHeight, info.height);
    }

    // 指定があれば基準枠を消去
    if (removeFrame) {
        frameInfo.obj.remove();
    }
}

main();