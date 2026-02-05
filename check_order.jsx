// 選択しテキストオブジェクトが連番になっているかのチェック
(function () {
    function toHalfWidth(str) {
        // 全角英数字を半角に変換
        str.contents = str.contents.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
            return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
        });
        return str;
    }
    function classifyChar(str) {
        if (/[0-9]/.test(str)) {
            return "suuji";
        } else if (/[A-Za-z]/.test(str)) {
            return "alphabet";
        } else if (/[\u30A0-\u30FF]/.test(str)) { // カタカナのUnicode範囲
            return "kana";
        } else {
            return "other";
        }
    }

    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません");
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    // テキストアイテムだけ抽出
    var textItems = [];
    for (var i = 0; i < selection.length; i++) {
        if (selection[i].typename === "TextFrame") {
            textItems.push(selection[i]);
        }
    }
    //全角を半角に
    for (var i = 0; i < textItems.length; i++) {
        textItems[i] = toHalfWidth(textItems[i]);
    }

    if (textItems.length === 0) {
        alert("テキストオブジェクトを選択してください");
        return;
    }

    var mojishubetu = classifyChar(textItems[0].contents.substr(0,1));
    if (mojishubetu == "other") {
        alert("選択されたテキストの文字種別が不正です。\r\n数字orアルファベットorカナを選択してください。");
        return;
    }

    // チェック方向を選択
    var dialog = new Window("dialog","マーク内連番チェックscript",[100,100,680,460],{closeButton:true,resizable:true}); // -> [^FootNote1]

    dialog.chuui2=dialog.add("statictext",{x:20, y:10, width:400, height:30},"※テキストの座標がバラバラの場合は正しくチェックできません");

    dialog.PnlHoukou = dialog.add('panel', {x:20, y:80, width:540, height:80}, "", {borderStyle:"inset"});
    dialog.PnlHoukou = dialog.PnlHoukou.add('panel', {x:10, y:5, width:510, height:60}, "【方向】");
    dialog.PnlHoukou.Z = dialog.PnlHoukou.add('radiobutton',{x:40, y:20, width:150, height:20},'Ｚ方向');
    dialog.PnlHoukou.N = dialog.PnlHoukou.add('radiobutton',{x:180, y:20, width:150, height:20},'Ｎ方向');

    dialog.PnlShubetu = dialog.add('panel', {x:20, y:175, width:540, height:80}, "", {borderStyle:"inset"});
    dialog.PnlShubetu = dialog.PnlShubetu.add('panel', {x:10, y:5, width:510, height:60}, "【文字種別】");
    dialog.PnlShubetu.Suuji = dialog.PnlShubetu.add('radiobutton',{x:40, y:20, width:150, height:20},'数字');
    dialog.PnlShubetu.Alfa = dialog.PnlShubetu.add('radiobutton',{x:180, y:20, width:150, height:20},'アルファベット');
    dialog.PnlShubetu.Kana = dialog.PnlShubetu.add('radiobutton',{x:370, y:20, width:150, height:20},'カタカナ');

    dialog.OkButton = dialog.add('button',{x:300, y:310, width:100, height:25},'OK');
    dialog.CancelButton = dialog.add('button',{x:410, y:310, width:100, height:25},'キャンセル');

    dialog.PnlHoukou.Z.value=true

    dialog.PnlShubetu.Suuji.value=true
    dialog.PnlShubetu.Suuji.enabled=true;
    dialog.PnlShubetu.Alfa.enabled=false;
    dialog.PnlShubetu.Kana.enabled=false;
    switch (mojishubetu){
    case "alphabet":
        dialog.PnlShubetu.Alfa.value=true;
        dialog.PnlShubetu.Suuji.enabled=false;
        dialog.PnlShubetu.Alfa.enabled=true;
        dialog.PnlShubetu.Kana.enabled=false;
        break;
    case "kana":
        dialog.PnlShubetu.Kana.value=true;
        dialog.PnlShubetu.Suuji.enabled=false;
        dialog.PnlShubetu.Alfa.enabled=false;
        dialog.PnlShubetu.Kana.enabled=true;
        break;
    }

    dialog.OkButton.onClick = function() {

        var direction = dialog.PnlHoukou.Z.value? "Z": "N";
        // var mojishubetu = dialog.PnlShubetu.Suuji.value? "suuji": (dialog.PnlShubetu.Alfa.value? "alphabet": "kana");

        dialog.close();

        // 配列を2次元に変換（縦列数と横列数の推定）
        function groupItems(items) {
            var cnt_1 = 0;//1次元の要素数
            var cnt_2 = 0;//2次元の要素数
            var arr = [];
            arr[cnt_1] = [];
            var item1 = 0;
            var item2 = 0;
            const epsilon = 0.002;//許容誤差
            for (var i = 0; i < items.length-1; i++) {
                arr[cnt_1].push(items[i]);
                cnt_2++;
                if (direction.toUpperCase() === "Z") {
                    item1 = items[i].top-items[i].height/2;
                    item2 = items[i+1].top-items[i+1].height/2;
                }
                else{
                    item1 = items[i].left+items[i].width/2;
                    item2 = items[i+1].left+items[i+1].width/2;
                }
                if (Math.abs(item1 - item2) > epsilon){
                    cnt_1++;
                    arr[cnt_1] = [];
                    cnt_2=0;
                }
            }
            arr[cnt_1].push(items[i]);
            return arr;
        }

        // 座標でソート（Z: 上→下, 左→右）（N: 左→右, 上→下）
        textItems.sort(function (a, b) {
            if (direction.toUpperCase() === "Z") {
                // 縦方向優先（上→下）, 同じyなら左→右
                if (Math.abs((a.top-a.height/2) - (b.top-b.height/2)) > 1){
                    return (b.top-b.height/2) - (a.top-a.height/2);
                }
                return (a.left+a.width/2) - (b.left+b.width/2);
            } else {
                // 横方向優先（左→右）, 同じxなら上→下
                if (Math.abs((a.left+a.width/2) - (b.left+b.width/2)) > 1){
                    return (a.left+a.width/2) - (b.left+b.width/2);
                }
                return (b.top-b.height/2) - (a.top-a.height/2);
            }
        });

        var grid = groupItems(textItems);
        var isValid = true;
        var item = 0;
        var ini = grid[0][0].contents[0];
        var arr_kana_han = ['ｱ','ｲ','ｳ','ｴ','ｵ','ｶ','ｷ','ｸ','ｹ','ｺ','ｻ','ｼ','ｽ','ｾ','ｿ','ﾀ','ﾁ','ﾂ','ﾃ','ﾄ','ﾅ','ﾆ','ﾇ','ﾈ','ﾉ','ﾊ','ﾋ','ﾌ','ﾍ','ﾎ','ﾏ','ﾐ','ﾑ','ﾒ','ﾓ','ﾔ','ﾕ','ﾖ','ﾗ','ﾘ','ﾙ','ﾚ','ﾛ','ﾜ','ｦ','ﾝ'];
        var arr_kana_zen = ['ア','イ','ウ','エ','オ','カ','キ','ク','ケ','コ','サ','シ','ス','セ','ソ','タ','チ','ツ','テ','ト','ナ','ニ','ヌ','ネ','ノ','ハ','ヒ','フ','ヘ','ホ','マ','ミ','ム','メ','モ','ヤ','ユ','ヨ','ラ','リ','ル','レ','ロ','ワ','ヲ','ン'];
        var kana_ini_INDEX = -1;
        var kana_arr;
        var iniCode;

        switch (mojishubetu){
        case "suuji":
            ini = parseInt(ini, 10);
            break;
        case "alphabet":
            ini = ini.toUpperCase();
            iniCode = ini.charCodeAt(0);
            break;
        case "kana":
            if (arr_kana_han.join().replace(/,/g,"").indexOf(ini)>0){
                kana_ini_INDEX = arr_kana_han.join().replace(/,/g,"").indexOf(ini);
                kana_arr = arr_kana_han;
            }
            else{
                kana_ini_INDEX = arr_kana_zen.join().replace(/,/g,"").indexOf(ini);
                kana_arr = arr_kana_zen;
            }
              break;
        }
        app.executeMenuCommand('deselectall')  //全オブジェクト選択解除
        for (var r = 0; r < grid.length; r++) {
            for (var c = 0; c < grid[r].length; c++) {
                var find_return = grid[r][c].contents.indexOf("\r");
                var expected;
                var value;
                if(find_return>-1){    //１文字ずつバラバラでない（改行区切り）
                    var result = true;
                    var arr = grid[r][c].contents.split(/\r/);
                    switch (mojishubetu){
                    case "suuji":

                        for (var cnt_current = 0; cnt_current < arr.length-1; cnt_current++) {
                            if (!(arr[cnt_current+1] - arr[cnt_current] == 1)){
                                result = false;
                            }
                        }
                        break;
                    case "alphabet":
                        for (var cnt_current = 0; cnt_current < arr.length-1; cnt_current++) {
                            if (!(arr[cnt_current+1].toUpperCase().charCodeAt(0) - arr[cnt_current].toUpperCase().charCodeAt(0) == 1)){
                                result = false;
                            }
                        }
                        break;
                    case "kana":
                        if (arr_kana_han.join().replace(/,/g,"").indexOf(arr[0])>0){
                            kana_arr = arr_kana_han;
                        }
                        else{
                            kana_arr = arr_kana_zen;
                        }
                        for (var cnt_current = 0; cnt_current < arr.length-1; cnt_current++) {
                            if (!(arr[cnt_current].toUpperCase() == kana_arr[cnt_current])){
                                result = false;
                            }
                        }
                        break;
                    }
                    if (!result) {
                        item=0;
                        for (var i = 0; i < r; i++) {
                            item=item+grid[i].length
                        }
                        textItems[item+c].selected = true;
                        isValid = false;
                    }
                }
                else{   //１文字ずつバラバラ
                    switch (mojishubetu){
                    case "suuji":
                        expected = parseInt(ini,10)+c;
                        value = parseInt(grid[r][c].contents.replace(/[^\d]/g, ""), 10);
                        break;
                    case "alphabet":
                        expected = String.fromCharCode(iniCode + c);
                        value = grid[r][c].contents.toUpperCase();
                        break;
                    case "kana":
                        expected = kana_arr[kana_ini_INDEX + c];
                        value = grid[r][c].contents.toUpperCase();
                        break;
                    }
                    if (value !== expected) {
                        item=0;
                        for (var i = 0; i < r; i++) {
                            item=item+grid[i].length
                        }
                        textItems[item+c].selected = true;
                        isValid = false;
                    }
                }
            }
        }

        if (isValid) {
            alert("連番チェックＯＫです");
        } else {
            alert("連番が崩れてるところがあります");
        }
    }

    //キャンセルボタン
    dialog.CancelButton.onClick = function() {
        dialog.close();
    }

    dialog.center();
    dialog.show();

})();