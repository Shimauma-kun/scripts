(function () {
sort_main();

function sort_main(){

    // ダイアログを表示してソート方向を選択
    var dialog = new Window("dialog", "ソート方向選択");
    
    dialog.alignChildren = "left";
    
    var radioGroup = dialog.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";
    
    var radioN = radioGroup.add("radiobutton", undefined, "N方向でソート");
    var radioZ = radioGroup.add("radiobutton", undefined, "Z方向でソート");
    
    // デフォルトはN方向を選択
    radioN.value = true;
    
    var buttonGroup = dialog.add("group");
    buttonGroup.orientation = "row";
    buttonGroup.alignment = "center";
    
    var okButton = buttonGroup.add("button", undefined, "OK", {name: "ok"});
    var cancelButton = buttonGroup.add("button", undefined, "キャンセル", {name: "cancel"});
    
    if (dialog.show() == 2) {
        // キャンセルボタンが押された場合は終了
        return;
    }
    
    var selitems = app.activeDocument.selection;

    //並び替え
    sortPosition(selitems,radioN.value);
    for(var i = 0; i < selitems.length; i++){
        selitems[i].name = selitems[i].name + "_" +i;
    }
    sortLayer();
    for(var i = 0; i < selitems.length; i++){
        selitems[i].name = selitems[i].name.replace(/_\d+$/, "");
    }
}

function sortLayer() {
    Array.prototype.forEach = function (callback) {
        for (var i = 0; i < this.length; i++)
            callback(this[i], i, this);
    };
    var layer = app.activeDocument.activeLayer;
    var items = []
    for (var i = 0; i < layer.pageItems.length; i++)
        items.push(layer.pageItems[i])
        items.sort(function (a, b) {
        if (rGetNumLen(a.name) === rGetNumLen(b.name)) return rGetNum(a.name).localeCompare(rGetNum(b.name))
        else if (rGetNumLen(a.name) < rGetNumLen(b.name)) return -1
        else if (rGetNumLen(a.name) > rGetNumLen(b.name)) return 1;
    })
    items.forEach(function (item) {
        item.move(layer, ElementPlacement.PLACEATEND)
    })
}

function sortPosition(r,order){
    var hs = [];
    var vs = [];
    for(var i = 0, iEnd = r.length; i < iEnd; i++){
        hs.push(r[i].left);
        vs.push(r[i].top);
    }
    if(order){
        //N方向
        r.sort(function(a,b){ return compPosition(a.left, b.left, b.top, a.top) });
    } else {
        //Z方向
        r.sort(function(a,b){ return compPosition(b.top, a.top, a.left, b.left) });
    }
}
function compPosition(a1, b1, a2, b2){
    return a1 == b1 ? a2 - b2 : a1 - b1;
}
function rMax(r){
    return Math.max.apply(null, r);
}
function rMin(r){
    return Math.min.apply(null, r);
}

function rGetNum(r){
    return r.toString().replace(/.*_/, '');
}
function rGetNumLen(r){
    return r.toString().replace(/.*_/, '').length;
}

})();
