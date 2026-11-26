#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("请先打开一个文档。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert("请先选中需要粘贴的对象。");
        return;
    }

    // 创建选择弹窗
    var dlg = new Window("dialog", "批量粘贴至画板");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 16;

    var panel = dlg.add("panel", undefined, "目标画板范围");
    panel.orientation = "column";
    panel.alignChildren = ["left", "top"];
    panel.margins = [14, 18, 14, 14];

    var rbOdd = panel.add("radiobutton", undefined, "仅奇数页 (1, 3, 5...)");
    var rbEven = panel.add("radiobutton", undefined, "仅偶数页 (2, 4, 6...)");
    rbOdd.value = true; // 默认选中奇数页

    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["right", "top"];
    var cancelBtn = btnGroup.add("button", undefined, "取消", { name: "cancel" });
    var okBtn = btnGroup.add("button", undefined, "确定", { name: "ok" });

    if (dlg.show() !== 1) return;

    var activeIdx = doc.artboards.getActiveArtboardIndex();
    var sourceRect = doc.artboards[activeIdx].artboardRect; // [left, top, right, bottom]
    var startIdx = rbOdd.value ? 0 : 1; // 索引 0 为第 1 页（奇数），索引 1 为第 2 页（偶数）

    for (var i = startIdx; i < doc.artboards.length; i += 2) {
        // 跳过原对象所在的画板，避免原地重复叠加
        if (i === activeIdx) continue;

        var targetRect = doc.artboards[i].artboardRect;
        var dx = targetRect[0] - sourceRect[0];
        var dy = targetRect[1] - sourceRect[1];

        // 逆序复制以保持多选元素的原有上下层级顺序
        for (var j = sel.length - 1; j >= 0; j--) {
            var dup = sel[j].duplicate();
            dup.position = [sel[j].position[0] + dx, sel[j].position[1] + dy];
        }
    }
})();