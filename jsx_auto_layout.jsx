#target illustrator

// オブジェクトの「見た目上の本当のバウンディングボックス」を深堀りして取得する機能
function getTrueBounds(obj) {
    if (obj.typename === "GroupItem") {
        if (obj.clipped) {
            // 【修正ポイント】
            // 複合パスには clipping 属性が存在しない場合があるため、
            // 「クリッピンググループ内の先頭（一番上）のアイテムが必ずマスクである」という仕様を利用する
            if (obj.pageItems.length > 0) {
                return obj.pageItems[0].geometricBounds;
            }
        } else {
            // 通常のグループの場合、さらに奥の中身をすべてチェックしてサイズを合成する
            var L = Infinity, T = -Infinity, R = -Infinity, B = Infinity;
            var hasValidBounds = false;
            for (var i = 0; i < obj.pageItems.length; i++) {
                var b = getTrueBounds(obj.pageItems[i]); // 再帰処理（マトリョーシカを開ける）
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
    
    // グループ以外（通常のパス、複合パス、テキストなど）の場合
    try {
        return obj.geometricBounds;
    } catch(e) {
        return null;
    }
}

// 取得したバウンディングボックスから中心座標を計算する
function getTrueCenter(obj) {
    var bounds = getTrueBounds(obj);
    if (!bounds) {
        bounds = obj.geometricBounds; // 万が一のための保険
    }
    var centerX = (bounds[0] + bounds[2]) / 2;
    var centerY = (bounds[1] + bounds[3]) / 2;
    return {x: centerX, y: centerY};
}

function alignToFrames() {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length < 2) {
        alert("エラー：中身のパスと、配置先の枠を一緒に選択してから実行してください。");
        return;
    }

    // 一番手前にあるものを「中身」として認識
    var contentObj = sel[0];

    // 残りのオブジェクトを「枠」として処理
    for (var i = 1; i < sel.length; i++) {
        var frame = sel[i];

        // 枠の「本当の中心」を取得
        var frameCenter = getTrueCenter(frame);

        // 中身を複製
        var newObj = contentObj.duplicate();

        // 複製した中身の「本当の中心」を取得
        var newObjCenter = getTrueCenter(newObj);

        // 差分を計算して移動
        var deltaX = frameCenter.x - newObjCenter.x;
        var deltaY = frameCenter.y - newObjCenter.y;
        newObj.translate(deltaX, deltaY);
    }

    // 配置元のオリジナルオブジェクトを削除
    contentObj.remove();
}

alignToFrames();