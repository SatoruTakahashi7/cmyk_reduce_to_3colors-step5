/**
 * CMYK4色を "C/M/Yのうち1チャネルを0にし、残り2チャネル + K" で近似するサンプルスクリプト.
 *
 * 【主な仕様】
 * - 選択中のオブジェクト（PathItem/GroupItem/CompoundPathItem/TextFrame）の塗り・線が
 *   CMYKカラーの場合に変換。
 * - CMYK→Lab 変換には app.convertSampleColor() を使用（Illustratorのカラー設定で
 *   "Japan Color 2001 Coated"が選ばれている前提）。
 * - (C, M, Y) のうち1チャネルを0にして2チャネル + Kの計3チャネルで総当たりし、
 *   Lab色差(ΔE76)が最小となるものを求める。
 * - 変換結果は最後に Math.round() で整数に丸める。
 * - 計算負荷が高いため、STEP値が小さいと処理が非常に遅くなる可能性があります。
 *
 * 【使い方】
 * 1) Illustratorで作業用CMYKプロファイルを"Japan Color 2001 Coated"に設定しておく。
 * 2) 変換したいオブジェクト（CMYK塗り/線を持つもの）を選択。
 * 3) このスクリプト(.jsx)を実行すると、選択オブジェクトのCMYKカラーが3色化される。
 * 4) しかし、浮動小数表現の誤差が完全にゼロにならないので、数値を手入力し直さないとならない（どうすればいいのかわからない）。
 */

(function () {
    try {
        //---- (A) 前準備・設定 ----
        if (!app.documents.length) {
            alert("ドキュメントが開いていません。");
            return;
        }
        var doc = app.activeDocument;

        var sel = doc.selection;
        if (!sel || sel.length < 1) {
            alert("オブジェクトが選択されていません。");
            return;
        }

        // ループのステップ幅 (小さくするほど計算が重い)
        var STEP = 5;

        //---- (B) カラー変換ユーティリティ ----

        /**
         * CMYK → Lab 変換（app.convertSampleColor使用）
         * @param {number} c 0~100
         * @param {number} m 0~100
         * @param {number} y 0~100
         * @param {number} k 0~100
         * @returns {number[]} [L, a, b]
         */
        function cmykToLab(c, m, y, k) {
            var cmykArr = [c, m, y, k];
            var labArr = app.convertSampleColor(
                ImageColorSpace.CMYK,
                cmykArr,
                ImageColorSpace.LAB,
                ColorConvertPurpose.defaultpurpose
            );
            return labArr; // [L,a,b]
        }

        /**
         * Lab同士のΔE76を計算
         * @param {number[]} lab1 [L,a,b]
         * @param {number[]} lab2 [L,a,b]
         * @returns {number} ΔE76
         */
        function deltaE76(lab1, lab2) {
            var dL = lab1[0] - lab2[0];
            var da = lab1[1] - lab2[1];
            var db = lab1[2] - lab2[2];
            return Math.sqrt(dL * dL + da * da + db * db);
        }

        /**
         * [Lab] に最も近い「C/M/Yのうち1チャネルを0にして、残り2チャネル+K」(=3チャネル)を探す
         * STEP刻みで総当たり → ΔE最小を返す
         * @param {number[]} targetLab [L,a,b]
         * @param {number} step  (例:5)
         * @returns {{c: number, m: number, y: number, k: number, diff: number}}
         */
        function findBest3ColorByLab(targetLab, step) {
            var bestResult = { c:0, m:0, y:0, k:0, diff:999999 };

            // (A) ペア=C/M, Y=0
            for (var cC = 0; cC <= 100; cC += step) {
                for (var cM = 0; cM <= 100; cM += step) {
                    for (var cK = 0; cK <= 100; cK += step) {
                        var labTest = cmykToLab(cC, cM, 0, cK);
                        var dE = deltaE76(targetLab, labTest);
                        if (dE < bestResult.diff) {
                            bestResult.c = cC;
                            bestResult.m = cM;
                            bestResult.y = 0;
                            bestResult.k = cK;
                            bestResult.diff = dE;
                        }
                    }
                }
            }

            // (B) ペア=C/Y, M=0
            for (var cC = 0; cC <= 100; cC += step) {
                for (var cY = 0; cY <= 100; cY += step) {
                    for (var cK = 0; cK <= 100; cK += step) {
                        var labTest = cmykToLab(cC, 0, cY, cK);
                        var dE = deltaE76(targetLab, labTest);
                        if (dE < bestResult.diff) {
                            bestResult.c = cC;
                            bestResult.m = 0;
                            bestResult.y = cY;
                            bestResult.k = cK;
                            bestResult.diff = dE;
                        }
                    }
                }
            }

            // (C) ペア=M/Y, C=0
            for (var cM = 0; cM <= 100; cM += step) {
                for (var cY = 0; cY <= 100; cY += step) {
                    for (var cK = 0; cK <= 100; cK += step) {
                        var labTest = cmykToLab(0, cM, cY, cK);
                        var dE = deltaE76(targetLab, labTest);
                        if (dE < bestResult.diff) {
                            bestResult.c = 0;
                            bestResult.m = cM;
                            bestResult.y = cY;
                            bestResult.k = cK;
                            bestResult.diff = dE;
                        }
                    }
                }
            }

            // --- ここで最終的に "完全整数" をめざす ---
            // 結果を "文字列に丸め" → "数値に変換" することで、
            // 内部的にも 15.0 などの形式に固定しやすい。
            bestResult.c = parseFloat(bestResult.c.toFixed(0));
            bestResult.m = parseFloat(bestResult.m.toFixed(0));
            bestResult.y = parseFloat(bestResult.y.toFixed(0));
            bestResult.k = parseFloat(bestResult.k.toFixed(0));

            return bestResult;

        }

        //---- (C) オブジェクトのCMYKを3色化する処理 ----

        /**
         * 塗り・線がCMYKColorの場合に「CMYK4色 → 3色化」する
         * @param {*} art PathItem, TextFrame など
         */
        function convertArtworkColor(art) {
            // (1) 塗り
            if (art.fillColor && art.fillColor.typename === "CMYKColor") {
                var c = art.fillColor.cyan;
                var m = art.fillColor.magenta;
                var y = art.fillColor.yellow;
                var k = art.fillColor.black;

                var oldLab = cmykToLab(c, m, y, k);
                var best3 = findBest3ColorByLab(oldLab, STEP);

                var newColor = new CMYKColor();
                newColor.cyan    = best3.c;
                newColor.magenta = best3.m;
                newColor.yellow  = best3.y;
                newColor.black   = best3.k;
                art.fillColor = newColor;
            }

            // (2) 線
            if (art.strokeColor && art.strokeColor.typename === "CMYKColor") {
                var c2 = art.strokeColor.cyan;
                var m2 = art.strokeColor.magenta;
                var y2 = art.strokeColor.yellow;
                var k2 = art.strokeColor.black;

                var oldLab2 = cmykToLab(c2, m2, y2, k2);
                var best3_2 = findBest3ColorByLab(oldLab2, STEP);

                var newColor2 = new CMYKColor();
                newColor2.cyan    = best3_2.c;
                newColor2.magenta = best3_2.m;
                newColor2.yellow  = best3_2.y;
                newColor2.black   = best3_2.k;
                art.strokeColor = newColor2;
            }
        }

        /**
         * 選択オブジェクトを再帰的に巡回して CMYKカラーを変換
         * @param {*} items 
         */
        function traverseSelection(items) {
            for (var i = 0; i < items.length; i++) {
                var it = items[i];
                if (it.typename === "PathItem") {
                    convertArtworkColor(it);

                } else if (it.typename === "CompoundPathItem") {
                    for (var j = 0; j < it.pathItems.length; j++) {
                        convertArtworkColor(it.pathItems[j]);
                    }

                } else if (it.typename === "GroupItem") {
                    traverseSelection(it.pageItems);

                } else if (it.typename === "TextFrame") {
                    // TextFrame全体の fill/strokeColor のみ変換
                    // (文字単位のカラーは別途対応が必要)
                    convertArtworkColor(it);
                }
                // その他はスキップ
            }
        }

        //---- (D) 実行 ----
        traverseSelection(sel);

        alert("変換完了 (STEP=" + STEP + ")。CMYK値は整数に丸めています。");
    } catch(e) {
        // 例外があれば行番号付きで表示
        alert("エラー発生: " + e + "\nline: " + (e.line || "N/A"));
    }
})();
