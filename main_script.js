const zipFile = document.getElementById("zip-file");
const fileSelect = document.getElementById("file-select");
const dropArea = document.getElementById("drop-area");
const viewArea = document.getElementById("xml-viewarea");
const saveButton = document.getElementById("save-button");

let mainXml, styleXml, footnoteXml, endnoteXml, relsXml, numberingXml;
let documentFileNameWithoutExtension;
let imageFilesContentArray = Array();
let imageFilesNameArray = Array();

let isGraphicxUsed = false,
  isAmsMathUsed = false,
  isMultiRowUsed = false,
  isCancelUsed = false;

// ZIPファイル（.docxファイル）を選択
zipFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;

  // ZIPファイル（.docxファイル）を読み込む
  readFile(file);
  zipFile.value = null; // 同じファイルを連続して選択しても処理が行われるようにするため
});

// ボタンを押したときに↑にイベントを渡す
fileSelect.addEventListener("click", () => {
  if (zipFile) {
    zipFile.click();
  }
});

// ドラッグ＆ドロップ
dropArea.addEventListener("dragover", (event) => {
  event.preventDefault();
  event.stopPropagation();
  dropArea.classList.add("active");
});

dropArea.addEventListener("dragleave", (event) => {
  event.preventDefault();
  event.stopPropagation();
  dropArea.classList.remove("active");
});

dropArea.addEventListener("drop", (event) => {
  event.preventDefault();
  event.stopPropagation();
  dropArea.classList.remove("active");

  const files = event.dataTransfer.files;
  if (!files.length) return;

  const file = files[0];
  if (!file) return;

  // ZIPファイル（.docxファイル）を読み込む
  readFile(file);
});

// 内容をZIPファイルとして保存
saveButton.addEventListener("click", () => {
  // 出力用ZIPファイルを作成
  const zip = new JSZip();

  // .texファイルをZIP書庫に入れる
  zip.file(documentFileNameWithoutExtension + ".tex", viewArea.textContent);

  // 保存しておいた画像ファイルをZIP書庫に入れる
  for (let i = 0; i < imageFilesContentArray.length; i++) {
    zip.file("images/" + imageFilesNameArray[i], imageFilesContentArray[i]);
  }

  // ZIPファイルをダウンロード
  zip.generateAsync({ type: "blob" }).then((blob) => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = documentFileNameWithoutExtension + ".zip";
    a.click();
    window.URL.revokeObjectURL(url);
  });
});

// .docxファイルを読み込むメイン関数
async function readFile(docxFile) {
  // 開始時間を記録
  const startTime = Date.now();

  const headingTags = [
    "\\part",
    "\\chapter",
    "\\section",
    "\\subsection",
    "\\subsubsection",
    "\\paragraph",
    "\\subparagraph",
  ];

  // 出力ファイル用に、ファイル名から拡張子を削除し、半角スペースをアンダーバーに変換したものを準備
  documentFileNameWithoutExtension = docxFile.name.split(".").slice(0, -1).join(".").replace(/ /g, "_");

  // .docxファイルをZIPファイルとして読み込む
  let zip;
  try {
    zip = await JSZip.loadAsync(docxFile);
  } catch {
    alert("指定されたファイルはWordの.docxファイルではありません（エラー1）。");
    return;
  }

  // ZIPファイル内のファイル
  let mainXmlDoc, styleXmlDoc, footnoteXmlDoc, endnoteXmlDoc, relsXmlDoc, numberingXmlDoc;

  // ファイル一覧を取得
  zip.forEach(async function (relativePath, zipEntry) {
    switch (zipEntry.name) {
      case "word/document.xml":
        mainXmlDoc = zip.file(zipEntry.name);
        break;
      case "word/styles.xml":
        styleXmlDoc = zip.file(zipEntry.name);
        break;
      case "word/footnotes.xml":
        footnoteXmlDoc = zip.file(zipEntry.name);
        break;
      case "word/endnotes.xml":
        endnoteXmlDoc = zip.file(zipEntry.name);
        break;
      case "word/numbering.xml":
        numberingXmlDoc = zip.file(zipEntry.name);
        break;
      case "word/_rels/document.xml.rels":
        relsXmlDoc = zip.file(zipEntry.name);
        break;
      default:
        // 画像の保存
        if (zipEntry.name.startsWith("word/media/")) {
          let imageFileName = zipEntry.name.replace("word/media/", "");
          await getImageFile(zip.file(zipEntry.name), imageFileName);
        }
        //else{console.log('未処理: ' + zipEntry.name)};
        break;
    }
  });

  // XMLをパース
  mainXml = await getXmlRoot(mainXmlDoc);
  styleXml = await getXmlRoot(styleXmlDoc);
  footnoteXml = await getXmlRoot(footnoteXmlDoc);
  endnoteXml = await getXmlRoot(endnoteXmlDoc);
  numberingXml = await getXmlRoot(numberingXmlDoc);
  relsXml = await getXmlRoot(relsXmlDoc);

  // document.xmlが読み込まれたかどうかチェック
  if (mainXml === undefined) {
    alert("指定されたファイルは.docxファイルではありません。");
    return;
  }

  // mainXmlのルートにw:documentが入っているかチェック
  if (mainXml.nodeName !== "w:document") {
    alert("指定されたファイルのdocument.xmlが破損しています。");
    return;
  }

  //w:document -> w:bodyの直下の要素を取得
  const paragraphs = mainXml.childNodes[0].childNodes;

  //console.log('パラグラフの数:'+paragraphs.length);

  let outputText = new Array();
  outputText.push("\\preamble");

  for (let i = 0; i < paragraphs.length; i++) {
    switch (paragraphs[i].nodeName) {
      case "w:p":
        //パラグラフを処理
        const paragraphString = processParagraph(paragraphs[i]);

        //画像とキャプションの処理
        const results_graphics = paragraphString.match(/\\includegraphics\[(.*?)\]{(.*?)}/);
        const results_caption = paragraphString.match(/\\caption{(.*?)}/);

        // パラグラフ内に画像とキャプションがある場合の処理
        if (results_graphics !== null && results_caption !== null) {
          const tmpstr = paragraphString
            .replace(/\\includegraphics\[(.*?)\]{(.*?)}/g, "")
            .replace(/\\caption{(.*?)}/g, "")
            .trim();
          if (tmpstr !== "") outputText.push(tmpstr);
          outputText.push("\\includegraphics[" + results_graphics[1] + "]{" + results_graphics[2] + "}");
          outputText.push("\\caption{" + results_caption[1] + "}");
          break;
        }

        // パラグラフ内に画像がある場合の処理
        if (results_graphics !== null) {
          const tmpstr = paragraphString.replace(/\\includegraphics\[(.*?)\]{(.*?)}/g, "").trim();
          if (tmpstr !== "") outputText.push(tmpstr);
          outputText.push("\\includegraphics[" + results_graphics[1] + "]{" + results_graphics[2] + "}");
          break;
        }

        // パラグラフ内にキャプションがある場合の処理
        if (results_caption !== null) {
          const tmpstr = paragraphString.replace(/\\caption{(.*?)}/g, "").trim();
          if (tmpstr !== "") outputText.push(tmpstr);
          outputText.push("\\caption{" + results_caption[1] + "}");
          break;
        }

        if (paragraphString !== "\\meaninglessparagraph") outputText.push(paragraphString);
        break;
      case "w:tbl":
        //テーブル
        const tableString = processTable(paragraphs[i]);
        if (tableString !== "\\meaninglessparagraph") outputText.push(tableString);
        break;
      default:
        //その他（w:bookmark, w:sectPrなどが該当するが、処理不要と思われる）
        //console.log('other');
        //console.log(paragraphs[i]);
        break;
    }
  }

  //最終行
  outputText.push("\\end{document}");

  //LaTeX形式を整える

  let nTitleLine = 0;
  let nSubtitleLine = 0;

  for (let i = 1; i < outputText.length; i++) {
    //リスト段落の処理
    if (outputText[i] !== null && outputText[i].startsWith("\\item")) {
      //リストの最適化（無駄な深いネストを浅くするために、ネストの最小値を調べる）
      let j;
      let mostShallowIndentLevel = 100;
      let listTypeString = Array(10).fill("itemize");

      for (j = 0; j < outputText.length - i; j++) {
        if (outputText[i + j].startsWith("\\item")) {
          const results = outputText[i + j].match(/\\item_([0-9]+)_([0-9])/);
          const indentLevel = parseInt(results[1]);
          const listType = parseInt(results[2]);
          if (mostShallowIndentLevel > indentLevel) mostShallowIndentLevel = indentLevel;
          listTypeString[indentLevel] = listType === 1 ? "enumerate" : "itemize";
        } else {
          break;
        }
      }

      let previousLevel = -1;
      let originalLevel = 0;
      let listString = "";

      //リストの最終処理
      for (let k = 0; k < j; k++) {
        listString = "";

        const results = outputText[i + k].match(/\\item_([0-9]+)_([0-9])/);
        originalLevel = parseInt(results[1]);
        let lvl = originalLevel - mostShallowIndentLevel;
        let itemText = outputText[i + k].replace(/\\item_([0-9]+)_([0-9])/, "");

        let lvldiff = lvl - previousLevel;

        if (lvldiff > 0) {
          //ネストが深くなる
          if (lvldiff == 1) {
            listString += "\\begin{" + listTypeString[originalLevel] + "}\n\\item " + itemText;
          } else {
            for (let l = 0; l < lvldiff; l++) {
              listString += "\\begin{" + listTypeString[originalLevel + l - 1] + "}\n\\item\n";
            }
            listString += itemText;
          }
        } else if (lvldiff < 0) {
          //ネストが浅くなる
          for (let l = 0; l > lvldiff; l--) {
            listString += "\\end{" + listTypeString[originalLevel + l + 1] + "}\n";
          }
          listString += "\\item " + itemText;
        } //同じネストの深さ＝アイテムの並列
        else {
          listString += "\\item " + itemText;
        }

        previousLevel = lvl;

        outputText[i + k] = listString.trim();
      }

      listString = "\n";

      for (let m = 0; m < previousLevel + 1; m++) {
        listString += "\\end{" + listTypeString[originalLevel - m] + "}\n";
      }
      outputText[i + j - 1] += listString;

      i += j - 1;
    }

    //図表のキャプションの処理
    else if (outputText[i] !== null && outputText[i].startsWith("\\caption{")) {
      let captionText = outputText[i].slice(9);
      captionText = captionText.trim().slice(0, -1);

      //図表のキャプションの不要な文字列を消去 https://office-watch.com/2022/all-named-format-switches-word-field-codes/
      captionText = RemoveCaptionHeader(captionText);

      if (outputText[i - 1].startsWith("\\[")) {
        if (i > 2 && outputText[i - 2].startsWith("\\[")) {
          //2つの数式に1つのキャプションがついている場合
          outputText[i] = outputText[i - 1].trim();
          outputText[i - 1] = outputText[i - 2].trim();
          outputText[i - 2] = "\\begin{figure}[hbtp]\n\\centering\n\\begin{minipage}<y>{\\linewidth}";
          outputText[i] += "\\caption{" + captionText + "}\n%\\label{}\n\\end{minipage}\n\\end{figure}\n";
        } else {
          outputText[i] = outputText[i - 1];
          outputText[i - 1] = "\\begin{figure}[hbtp]\n\\centering\n\\begin{minipage}<y>{\\linewidth}";
          outputText[i] += "\\caption{" + captionText + "}\n%\\label{}\n\\end{minipage}\n\\end{figure}\n";
        }
      } else if (outputText[i - 1].startsWith("\\includegraphics")) {
        outputText[i] = outputText[i - 1];
        outputText[i - 1] = "\\begin{figure}[hbtp]\n\\centering\n\\begin{minipage}<y>{\\linewidth}\n\\centering";
        outputText[i] += "\n\\caption{" + captionText + "}\n%\\label{}\n\\end{minipage}\n\\end{figure}\n";
      } else if (outputText[i - 1].startsWith("\\begin{tabular}")) {
        outputText[i] = outputText[i - 1];
        outputText[i - 1] =
          "\\begin{table}[hbtp]\n\\begin{minipage}<y>{\\linewidth}\n\\caption{" +
          captionText +
          "}\n" +
          "%\\label{}\n\\centering";
        outputText[i] += "\\end{minipage}\n\\end{table}\n";
        const regex = /\\addtocounter\s*([\s\S]*)\n\\end{minipage}\n\\end{table}/g;
        outputText[i] = outputText[i].replace(regex, "\\end{minipage}\n\\end{table}\n\\addtocounter$1");
      } else if (outputText[i + 1].startsWith("\\begin{tabular}")) {
        outputText[i] =
          "\\begin{table}[hbtp]\n\\begin{minipage}<y>{\\linewidth}\n\\caption{" +
          captionText +
          "}\n" +
          "%\\label{}\n\\centering";
        outputText[i + 1] += "\\end{minipage}\n\\end{table}\n";
        const regex = /\\addtocounter\s*([\s\S]*)\n\\end{minipage}\n\\end{table}/g;
        outputText[i + 1] = outputText[i + 1].replace(regex, "\\end{minipage}\n\\end{table}\n\\addtocounter$1");
      } else if (outputText[i + 1].startsWith("\\includegraphics")) {
        outputText[i] =
          "\\begin{figure}[hbtp]\n\\centering\n\\begin{minipage}<y>{\\linewidth}\n\\centering\n\\caption{" +
          captionText +
          "}";
        outputText[i + 1] += "%\\label{}\n\\end{minipage}\n\\end{figure}\n";
      } else if (outputText[i + 1].startsWith("\\[")) {
        if (i < nLine - 2 && outputText[i + 2].startsWith("\\[")) {
          //2つの数式に1つのキャプションがついている場合
          outputText[i] =
            "\\begin{figure}[hbtp]\n\\centering\n\\begin{minipage}<y>{\\linewidth}\n\\caption{" + captionText + "}";
          outputText[i + 1] = outputText[i + 1].trim();
          outputText[i + 2] += "%\\label{}\n\\end{minipage}\n\\end{figure}\n";
        } else {
          outputText[i] =
            "\\begin{figure}[hbtp]\n\\centering\n\\begin{minipage}<y>{\\linewidth}\n\\caption{" + captionText + "}";
          outputText[i + 1] += "%\\label{}\n\\end{minipage}\n\\end{figure}\n";
        }
      }
    }

    //見出し内脚注とサブタイトルの処理
    else if (outputText[i] !== null && outputText[i].startsWith("\\h")) {
      //複数脚注への対応
      let footnotePostfix = "",
        endnotePostfix = "";
      let footnoteCount = 0,
        endnoteCount = 0;

      //paragraph/subparagraphはfootnote,endnoteのままでいいので処理しない
      if (!headingTags[parseInt(outputText[i].substr(2, 1)) - 1 + parseInt(defaultHeadingTag)].includes("paragraph")) {
        let noteText = inParenString(outputText[i], "note", "{", "}");
        while (noteText != "") {
          //footnoteかendnoteか判定（末尾の文字で）
          let footorend = outputText[i].substr(outputText[i].indexOf("note{" + noteText) - 1, 1);
          if (footorend == "t") {
            footnotePostfix += "\n\\addtocounter{footnote}{1}\\footnotetext{" + noteText + "}";
            outputText[i] =
              outputText[i].substr(0, outputText[i].indexOf("\\footnote{" + noteText) + 9) +
              "mark " +
              outputText[i].substr(outputText[i].indexOf("\\footnote{" + noteText) + 11 + noteText.length);
            footnoteCount++;
          } else if (footorend == "d") {
            endnotePostfix += "\n\\addtocounter{endnote}{1}\\endnotetext{" + noteText + "}";
            outputText[i] =
              outputText[i].substr(0, outputText[i].indexOf("\\endnote{" + noteText) + 8) +
              "mark " +
              outputText[i].substr(outputText[i].indexOf("\\endnote{" + noteText) + 10 + noteText.length);
            endnoteCount++;
          } else console.log("処理できないノート：" + outputText[i]);
          noteText = inParenString(outputText[i], "note", "{", "}");
        }

        if (footnoteCount > 0) footnotePostfix = "\n\\addtocounter{footnote}{-" + footnoteCount + "}" + footnotePostfix;
        if (endnoteCount > 0) endnotePostfix = "\n\\addtocounter{endnote}{-" + endnoteCount + "}" + endnotePostfix;
      }

      //サブタイトル処理 A--B   {A}{B}
      const afterheaderstr = outputText[i].replace(/\\h(\d)\{(.*)\}(.*)/, "$3");
      const headerstr = outputText[i].replace(/\\h(\d)\{(.*)\}(.*)/, "\\h$1");
      let strtitle = outputText[i].replace(/\\h(\d)\{(.*)――(.*)\}(.*)/, "$2");
      let strsubtitle = outputText[i].replace(/\\h(\d)\{(.*)――(.*)\}(.*)/, "$3");

      if (strtitle.trim() === outputText[i].trim()) {
        strtitle = outputText[i].replace(/\\h(\d)\{(.*)\}(.*)/, "$2");
        strsubtitle = "";
      }

      let strtitle_without_footmark = strtitle.replace(/\\footnotemark /g, "").replace(/\\endnotemark /g, "");
      let strsubtitle_without_footmark = strsubtitle.replace(/\\footnotemark /g, "").replace(/\\endnotemark /g, "");

      if (strtitle != strtitle_without_footmark) {
        if (strsubtitle != "") {
          outputText[i] =
            headerstr.trim() +
            "[" +
            strtitle_without_footmark.trim() +
            "]" +
            "{" +
            strtitle.trim() +
            "}" +
            "[" +
            strsubtitle.trim() +
            "]" +
            footnotePostfix +
            endnotePostfix +
            afterheaderstr;
        } else {
          outputText[i] =
            headerstr.trim() +
            "[" +
            strtitle_without_footmark.trim() +
            "]" +
            "{" +
            strtitle.trim() +
            "}" +
            footnotePostfix +
            endnotePostfix +
            afterheaderstr;
        }
      } else {
        if (strsubtitle != "") {
          outputText[i] =
            headerstr.trim() +
            "{" +
            strtitle.trim() +
            "}" +
            "[" +
            strsubtitle.trim() +
            "]" +
            footnotePostfix +
            endnotePostfix +
            afterheaderstr;
        } else {
          outputText[i] =
            headerstr.trim() + "{" + strtitle.trim() + "}" + footnotePostfix + endnotePostfix + afterheaderstr;
        }
      }

      //無駄な表記を削除
      outputText[i] = outputText[i].replace("\\addtocounter{footnote}{-1}\n\\addtocounter{footnote}{1}", "");
      outputText[i] = outputText[i].replace("\\addtocounter{endnote}{-1}\n\\addtocounter{endnote}{1}", "");

      //h1～h5に、part-chapter-sectionを割り振る
      let n = parseInt(outputText[i].substr(2, 1)) - 1;
      outputText[i] =
        headingTags[n + parseInt(defaultHeadingTag)] + (isAutoNumbering ? "" : "*") + outputText[i].slice(3);

      //目次をTOCに送る（見出しに自動で番号を振らない場合）
      if (!isAutoNumbering) {
        let titleline = strtitle_without_footmark.trim();
        if (strsubtitle !== "") titleline += "――" + strsubtitle_without_footmark.trim();
        outputText[i] += `\\phantomsection\n\\addcontentsline{toc}{${headingTags[n + parseInt(defaultHeadingTag)].slice(
          1
        )}}{${titleline}}\n`;
      }

      if (outputText[i].startsWith("\\chapter")) isBook = "book";
    } else if (outputText[i] !== null && outputText[i].startsWith("\\title")) {
      if (nTitleLine === 0) nTitleLine = i;
      else if (nTitleLine === i - 1) nSubtitleLine = i;
    } else if (outputText[i] !== null && outputText[i].startsWith("\\subtitle")) {
      if (nSubtitleLine === 0) {
        if (nTitleLine === 0) nTitleLine = i;
        else nSubtitleLine = i;
      }
    }

    if (outputText[i] !== null) outputText[i] = replaceDashCharacters(outputText[i]);
  }

  //プリアンブルの作成
  outputText[0] = "";

  outputText[0] +=
    "%Cloud LaTeXでコンパイルする場合は、「メニュー」→「LaTeXエンジン」→「lualatex」を選択してください。\n";

  //jlreqのオプション
  let docoption = "";
  if (isBook !== "") docoption += isBook + ",";
  if (isTate !== "") docoption += isTate + ",";
  if (isTwoColumn !== "") docoption += isTwoColumn + ",";
  if (paperSize !== "") docoption += paperSize + ",";
  if (fontSize !== "") docoption += "jafontsize=" + fontSize + ",";
  if (docoption.length > 0) docoption = docoption.slice(0, -1); // 末尾の「, 」を消去
  outputText[0] += "\\documentclass[" + docoption + "]{jlreq}\n\n";

  //ブロック引用の設定
  outputText[0] += "%ブロック引用の設定\n";
  outputText[0] += "\\jlreqsetup{quote_beforeafter_space=0.5\\baselineskip}%ブロック引用の前後を0.5行空ける\n";

  //脚注と後注の設定
  outputText[0] += "\n%脚注と後注の設定\n";
  outputText[0] += "\\jlreqsetup{endnote_counter=endnote}\\newcounter{endnote}%脚注と後注で別のカウンターを使う\n";

  if (isBook) {
    outputText[0] += "\\jlreqsetup{endnote_position={_chapter}}%後注は章の末尾に入れる\n";
  } else {
    outputText[0] += "\\jlreqsetup{endnote_position={_part}}%後注は部の末尾に入れる\n";
  }

  if (isTate) {
    outputText[0] += "\\renewcommand{\\thefootnote}{（\\tatechuyoko*{\\arabic{footnote}}）}%脚注の形式は（1）\n"; //
    outputText[0] += "\\renewcommand{\\theendnote}{［\\tatechuyoko*{\\arabic{endnote}}］}%後注の形式は［1］\n";
  } else {
    outputText[0] += "\\renewcommand{\\thefootnote}{（\\arabic{footnote}\\hbox{}）}%脚注の形式は（1）\n"; //
    outputText[0] += "\\renewcommand{\\theendnote}{*\\arabic{endnote}\\hbox{}}%後注の形式は*1\n";
  }

  //ルビの設定
  outputText[0] += "\n%ルビと圏点（傍点）の設定\n";
  outputText[0] += "\\usepackage{luatexja-ruby}\\ltjsetruby{kenten=﹅,size=0.5}%圏点記号とサイズの指定\n\n";

  //PDFの開き方向の指定
  outputText[0] += "%PDFの開き方向の指定\n";
  if (isTate) outputText[0] += "\\usepackage[pdfdirection=R2L,hidelinks]{hyperref}%縦書き文書は右開き\n";
  else outputText[0] += "\\usepackage[pdfdirection=L2R,hidelinks]{hyperref}%横書き文書は左開き\n";
  outputText[0] += "\\usepackage{bookmark}\n";

  //その他必要なパッケージの読み込み
  outputText[0] += "\\usepackage{lltjext}\n"; //<y>をするために必須
  if (useSmallFontSizeInBracket) {
    outputText[0] += "\\usepackage{relsize}\n";
  }
  if (isGraphicxUsed) outputText[0] += "\\usepackage{graphicx}\n";
  if (isAmsMathUsed) outputText[0] += "\\usepackage{amsmath}\n";
  if (isMultiRowUsed) outputText[0] += "\\usepackage{multirow}\n";
  if (isCancelUsed) outputText[0] += "\\usepackage{cancel}\n";
  //if (bUdline) outputText[0] += "\\usepackage{udline}\n"; //＊＊

  //独自マクロ
  outputText[0] += "\n%各種マクロ\n";
  outputText[0] +=
    '\\usepackage{newunicodechar}\\makeatletter\\chardef\\my@J@horizbar="2015\\newunicodechar{―}{\\x@my@dash}\\def\\x@my@dash{\\@ifnextchar―{\\my@J@horizbar\\kern-.5\\zw\\my@J@horizbar\\kern-.5\\zw}{\\my@J@horizbar}}\\makeatother%ダッシュをつなげる\n';
  outputText[0] +=
    '\\usepackage{newunicodechar}\\makeatletter\\chardef\\my@J@tdreader="2026\\newunicodechar{…}{\\x@my@tdrdef}\\def\\x@my@tdrdef{\\ifnum\\ltjgetparameter{direction}=3{︙}\\else \\my@J@tdreader\\fi}\\makeatother%縦書き時の三点リーダ\n';
  if (isFullWidthBracket) {
    outputText[0] +=
      "\\usepackage{relsize, newunicodechar}\\newif\\iffoot\\footfalse\\newcounter{parnest}\\setcounter{parnest}{0}%（）内の級数下げマクロ：変数の準備（級数下げしたくない箇所は、（）の前後を\\foottrue～\\footfalseで括ること）\n";
    outputText[0] +=
      "\\let\\origfootnotetext\\footnotetext\\renewcommand{\\footnotetext}[2][]{\\ifx\\relax#1\\relax \\origfootnotetext{\\foottrue #2 \\footfalse}\\else\\origfootnotetext[#1]{\\foottrue #2 \\footfalse}\\fi}\\let\\origfootnote\\footnote\\renewcommand{\\footnote}[1]{\\ifnum\\ltjgetparameter{direction}=3\\origfootnote{\\foottrue #1 \\footfalse}\\else\\origfootnote{\\foottrue #1 \\footfalse}\\fi}%脚注コマンドを変更\n";
    outputText[0] +=
      "\\let\\origendnotetext\\endnotetext\\renewcommand{\\endnotetext}[2][]{\\ifx\\relax#1\\relax \\origendnotetext{\\foottrue #2 \\footfalse}\\else\\origendnotetext[#1]{\\foottrue #2 \\footfalse}\\fi}\\let\\origendnote\\endnote\\renewcommand{\\endnote}[1]{\\ifnum\\ltjgetparameter{direction}=3\\origendnote{\\foottrue #1 \\footfalse}\\else\\origendnote{\\foottrue #1 \\footfalse}\\fi}%文末脚注コマンドを変更\n";
    outputText[0] +=
      '\\makeatletter\\chardef\\my@J@kakkostart="FF08\\newunicodechar{（}{\\iffoot\\my@J@kakkostart\\else\\addtocounter{parnest}{1}\\ifnum\\value{parnest}=1 \\relsize{-0.5}\\my@J@kakkostart\\else\\my@J@kakkostart\\fi\\fi}\\makeatother%開くカッコは脚注外では級数下げ\n';
    outputText[0] +=
      '\\makeatletter\\chardef\\my@J@kakkoend="FF09\\newunicodechar{）}{\\iffoot\\my@J@kakkoend\\else\\addtocounter{parnest}{-1}\\ifnum\\value{parnest}=0 \\my@J@kakkoend\\relsize{0.5}\\else\\my@J@kakkoend\\fi\\fi}\\makeatother%閉じるカッコは脚注外では級数上げ（元に戻す）\n';
    outputText[0] +=
      '\\makeatletter\\chardef\\my@J@kikkostart="3014\\newunicodechar{〔}{\\iffoot\\my@J@kikkostart\\else\\addtocounter{parnest}{1}\\ifnum\\value{parnest}=1 \\relsize{-0.5}\\my@J@kikkostart\\else\\my@J@kikkostart\\fi\\fi}\\makeatother%開く亀甲カッコは脚注外では級数下げ\n';
    outputText[0] +=
      '\\makeatletter\\chardef\\my@J@kikkoend="3015\\newunicodechar{〕}{\\iffoot\\my@J@kikkoend\\else\\addtocounter{parnest}{-1}\\ifnum\\value{parnest}=0 \\my@J@kikkoend\\relsize{0.5}\\else\\my@J@kikkoend\\fi\\fi}\\makeatother%閉じる亀甲カッコは脚注外では級数上げ（元に戻す）\n';
  }

  outputText[0] += "\n";

  //著者名と所属を取得
  let nLastTitleLine = nTitleLine;
  let strAuthorsLine = "";

  if (nLastTitleLine < nSubtitleLine) nLastTitleLine = nSubtitleLine;

  if (nLastTitleLine > 0) {
    for (let i = 1; i < 3; i++) {
      let tmpstr = outputText[nLastTitleLine + i].trim();
      tmpstr = tmpstr.replace(/\\footnote\{(.*?)\}/g, ""); // ?をつけると最短一致で検索 //バグ：ここで書式があったら…
      if (tmpstr.includes("\\")) break;
      if (tmpstr.includes("目次")) break;

      let separator = "";
      if (tmpstr.includes("　")) separator = "　";
      if (tmpstr.includes("，")) separator = "，";
      if (tmpstr.includes("、")) separator = "、";

      let bQuit = false;
      const names = tmpstr.split(separator);
      names.forEach(function (name_ind) {
        if (name_ind.length > 18) {
          //18文字より長ければ名前ではない
          bQuit = true;
        }
      });
      if (bQuit) break;

      if (separator !== "") {
        //著者が複数の場合
        const regex = new RegExp(separator, "g");
        tmpstr = outputText[nLastTitleLine + i].trim().replace(regex, "🔣");
        const names2 = tmpstr.split("🔣");
        names2.forEach(function (name2_ind) {
          if (name2_ind.length > 1) {
            // 「訳」とか「著」とかは名前ではない
            strAuthorsLine += name2_ind.replace("\\footnote", "\\thanks") + " \\and ";
          }
        });
        if (strAuthorsLine.endsWith(" \\and ")) strAuthorsLine = strAuthorsLine.slice(0, -6);
      } //著者がひとりの場合
      else {
        strAuthorsLine = outputText[nLastTitleLine + i].trim().replace("\\footnote", "\\thanks");
      }
      if (strAuthorsLine !== "") {
        outputText[nLastTitleLine + i] = "";
        strAuthorsLine = "\\author{" + strAuthorsLine + "}";
        break;
      }
    }
  }

  if (strAuthorsLine === "") strAuthorsLine = "\\author{著者名\\thanks{所属}}";

  //タイトルの処理
  if (nTitleLine > 0 && nSubtitleLine > 0) {
    outputText[nTitleLine] = outputText[nTitleLine].trim();
    outputText[nTitleLine] = outputText[nTitleLine].slice(0, -1);
    outputText[nTitleLine] += "\\large{\\\\";
    outputText[nSubtitleLine] = outputText[nSubtitleLine].trim();
    outputText[nSubtitleLine] = outputText[nSubtitleLine].replace("\\subtitle{", "").replace("\\title{", "");
    outputText[nSubtitleLine] = outputText[nSubtitleLine].slice(0, -1);
    outputText[nTitleLine] += outputText[nSubtitleLine] + "}}\n";
    outputText[nTitleLine] += strAuthorsLine + "\n";
    outputText[nTitleLine] += "\\date{}\n\n\\begin{document}\n\n\\maketitle";
    outputText[nSubtitleLine] = "";
    outputText[0] += outputText[nTitleLine];
    outputText[nTitleLine] = "";
  } else if (nTitleLine > 0) {
    outputText[nTitleLine] += strAuthorsLine + "\n";
    outputText[nTitleLine] += "\\date{}\n\n\\begin{document}\n\n\\maketitle";
    outputText[0] += outputText[nTitleLine];
    outputText[nTitleLine] = "";
  } else if (nSubtitleLine > 0) {
    outputText[nSubtitleLine] = outputText[nSubtitleLine].replace("\\subtitle{", "\\title{");
    outputText[nSubTitleLine] += strAuthorsLine + "\n";
    outputText[nSubtitleLine] += "\\date{}\n\n\\begin{document}\n\n\\maketitle";
    outputText[0] += outputText[nSubtitleLine];
    outputText[nSubtitleLine] = "";
  } else {
    outputText[0] += "\n\\begin{document}\n";
  }

  // 文書中に表題、副題が複数ある場合は見出しレベル1に変更する（＊この場合にTOCに飛ばない…）
  for (let i = 1; i < outputText.length; i++) {
    outputText[i] = outputText[i]
      .replace(/\\title/g, headingTags[parseInt(defaultHeadingTag)])
      .replace(/\\subtitle/g, headingTags[parseInt(defaultHeadingTag)]);
  }

  // TeXテキストを出力＋段落頭の\cancelの対策
  let texString = outputText.join("\n").replace(/\n\n\\cancel{/g, "\n\n\\hspace{0.01px}\\cancel{");

  if (isConvertEMFFile) {
    texString = texString.replace(/\\includegraphics\[(.*?)\]{(.*?).emf}/g, "\\includegraphics[$1]{$2.png}");
  } else {
    texString = texString.replace(
      /\\includegraphics\[(.*?)\]{(.*?).emf}/g,
      "%.emfファイルは手動で変換してください。\\includegraphics[$1]{$2.emf}"
    );
  }

  // 終了時間を記録し、経過時間を表示する
  const endTime = Date.now();
  viewArea.textContent = "%" + (endTime - startTime) / 1000 + "秒で変換しました。\n" + texString;

  // 保存ボタンを有効化
  saveButton.style.display = "block";

  // TeXのソースコードプレビューを表示
  const sourceCodeElement = document.querySelector(".sourcecode");
  sourceCodeElement.style.setProperty(
    "--content",
    '"' + "プレビュー: " + documentFileNameWithoutExtension + ".tex" + '"'
  );
  sourceCodeElement.style.display = "block";
}

// イメージファイルの保存
async function getImageFile(imageFile, imageFileName) {
  if (imageFileName.endsWith(".emf") && isConvertEMFFile) {
    imageFilesContentArray.push(await getPNGfromEMF(imageFile));
    imageFilesNameArray.push(imageFileName.slice(0, -4) + ".png");
  } else {
    imageFilesContentArray.push(await imageFile.async("uint8array"));
    imageFilesNameArray.push(imageFileName);
  }
}

// .emfファイルを.pngファイルに変換
async function getPNGfromEMF(imageFile) {
  var pNum = 0; // number of the page, that you want to render
  var scale = 1; // the scale of the document
  var wrt = new ToContext2D(pNum, scale);
  FromEMF.Parse(await imageFile.async("uint8array"), wrt);

  // Canvasの2Dコンテキストを取得
  const convertedCanvas = wrt.canvas;
  const ctx = convertedCanvas.getContext("2d");

  // 新しいCanvasを作成し、2Dコンテキストを取得
  const tempCanvas = document.createElement("canvas");
  tempCanvas.width = convertedCanvas.width;
  tempCanvas.height = convertedCanvas.height;
  const tempCtx = tempCanvas.getContext("2d");

  // 新しいCanvasで上下反転
  tempCtx.scale(1, -1);
  tempCtx.translate(0, -convertedCanvas.height);

  // 元のCanvasの内容を反転してコピー
  tempCtx.drawImage(convertedCanvas, 0, 0);

  // 反転したCanvasの内容をデータURIとして取得
  const dataURI = tempCanvas.toDataURL("image/png");

  // データURIからBlob形式のデータを作成
  const byteString = atob(dataURI.split(",")[1]);
  const ab = new ArrayBuffer(byteString.length);
  const ia = new Uint8Array(ab);
  for (let i = 0; i < byteString.length; i++) {
    ia[i] = byteString.charCodeAt(i);
  }
  return ia;
}

// 段落の処理
function processParagraph(paragraphXml, isInFootnote = false, isInTable = false) {
  // 再帰的処理の際に参照渡しする変数群をオブジェクトに格納しておく
  let state = {
    paragraphPrefix: "",
    paragraphPostfix: "",
    runPrefix: "",
    runPostfix: "",
    isMeaninglessParagraph: true,
    isMeaninglessRun: true,
    eastAsianID: "",
    fcString: "",
  };

  // 段落スタイルの取得（注やテーブルの中のものは無視する）
  if (!isInFootnote && !isInTable) {
    const paragraphProperties = getXElement(paragraphXml, "pPr");
    if (paragraphProperties !== undefined) {
      processParagraphProperties(paragraphProperties, state);
    }
  }

  //パラグラフ以下の要素の処理
  const paragraphItems = paragraphXml.childNodes;

  for (let i = 0; i < paragraphItems.length; i++) {
    state.isMeaninglessRun = true;
    state.runPrefix = "";
    state.runPostfix = "";

    //ハイパーリンク等の場合、一段階層が深くなる（改版履歴ON時の追記も）
    if (paragraphItems[i].nodeName === "w:hyperlink" || paragraphItems[i].nodeName === "w:ins") {
      //目次TOCなどは除外
      if (paragraphItems[i].getAttribute("w:anchor") === null) {
        const hyperlinkParagraphItems = paragraphItems[i].querySelectorAll(":scope > r");
        for (let j = 0; j < hyperlinkParagraphItems.length; j++) {
          //インライン数式:oMath
          if (hyperlinkParagraphItems[j].nodeName === "m:oMath") {
            console.log("hyperlink inline math");
            console.log(hyperlinkParagraphItems[j]);
            state.runPrefix += processMath(hyperlinkParagraphItems[j], false) + " ";
            state.isMeaninglessRun = false;
          }
          //パラグラフ数式:oMathPara
          if (hyperlinkParagraphItems[j].nodeName === "m:oMathPara") {
            console.log("hyperlink para math");
            console.log(hyperlinkParagraphItems[j]);
            state.paragraphPrefix += processMath(hyperlinkParagraphItems[j], true);
            state.isMeaninglessParagraph = false;
          }
          //文節:r
          if (hyperlinkParagraphItems[j].nodeName === "w:r") {
            processRun(hyperlinkParagraphItems[j], state);
          }
        }
      } else {
        //if (paragraphItems[i].getAttribute("w:anchor") === null) break; //＊＊これでいい？ループ抜けない？
        //if (paragraphItems[i].getAttribute("w:anchor").toLowerCase().startsWith("_toc")) {
        //目次TOC等の場合、特に何も処理しない
        //}
      }
    }

    //インライン数式:oMath
    if (paragraphItems[i].nodeName === "m:oMath") {
      state.runPrefix += processMath(paragraphItems[i], false) + " ";
      state.isMeaninglessRun = false;
    }

    //パラグラフ数式:oMathPara
    if (paragraphItems[i].nodeName === "m:oMathPara") {
      state.paragraphPrefix += processMath(paragraphItems[i], true);
      state.isMeaninglessParagraph = false;
    }

    //文節:r
    if (paragraphItems[i].nodeName === "w:r") {
      processRun(paragraphItems[i], state);
    }

    if (!state.isMeaninglessRun) {
      state.paragraphPrefix += state.runPrefix + state.runPostfix;
      state.isMeaninglessParagraph = false;
    } else {
      //console.log('meaninglessrun');
      //console.log(paragraphItems[i]);
    }
  }

  if (!isInTable) state.paragraphPostfix += "\n";
  if (isInFootnote) state.paragraphPostfix = state.paragraphPostfix.trimEnd() + " \\\\\\indent ";
  state.paragraphPrefix = state.paragraphPrefix.trimStart(); //段落頭の全角スペースをトル

  //fwb＊＊
  //marginpar＊＊

  if (!state.isMeaninglessParagraph) {
    let paragraphTextFinalized = (state.paragraphPrefix + state.paragraphPostfix)
      .replace(/\\kenten}\\kenten{/g, "") //重複タグを削除
      .replace(/\\fbox}\\fbox{/g, "")
      .replace(/\\Sl}\\Sl{/g, "")
      .replace(/\\ul}\\ul{/g, "")
      .replace(/\\textit}\\textit{/g, "")
      .replace(/\\textbf}\\textbf{/g, "")
      .replace(/\\textsuperscript}\\textsuperscript{/g, "")
      .replace(/\\textsubscript}\\textsubscript{/g, "")
      .replace(/\\kenten}/g, "}") //不要な表記を削除
      .replace(/\\fbox}/g, "}")
      .replace(/\\Sl}/g, "}")
      .replace(/\\ul}/g, "}")
      .replace(/\\textit}/g, "}")
      .replace(/\\textbf}/g, "}")
      .replace(/\\textsuperscript}/g, "}")
      .replace(/\\textsubscript}/g, "}");

    //脚注内にダッシュがあると見出しのサブタイトル処理が誤動作するため文字を変えておく
    if (isInFootnote) return paragraphTextFinalized.replace(/――/g, "──");
    else if (useSmallFontSizeInBracket && !paragraphTextFinalized.startsWith("\\h")) {
      if (
        countChar(paragraphTextFinalized, "（") === countChar(paragraphTextFinalized, "）") &&
        countChar(paragraphTextFinalized, "〔") === countChar(paragraphTextFinalized, "〕")
      ) {
        return reduceFontSizeWithinParentheses(paragraphTextFinalized);
      } else {
        return paragraphTextFinalized;
      }
    } else {
      return paragraphTextFinalized;
    }
  } else {
    return "\\meaninglessparagraph";
  }
}

// 段落ごとのスタイルを処理
function processParagraphProperties(paragraphProperties, state) {
  const styleTag = getXElement(paragraphProperties, "pStyle");
  if (styleTag === undefined) return;

  const styleID = styleTag.getAttribute("w:val");
  const styleElement = getXElementByAttribute(styleXml, "style", "w:styleId", styleID);
  const styleName = getXElement(styleElement, "name").getAttribute("w:val");

  switch (styleName.toLowerCase()) {
    case "title":
      state.paragraphPrefix = "\\title{";
      state.paragraphPostfix = "}";
      break;
    case "subtitle":
      state.paragraphPrefix = "\\subtitle{";
      state.paragraphPostfix = "}";
      break;
    case "heading 1":
      state.paragraphPrefix = "\\h1{";
      state.paragraphPostfix = "}";
      break;
    case "heading 2":
      state.paragraphPrefix = "\\h2{";
      state.paragraphPostfix = "}";
      break;
    case "heading 3":
      state.paragraphPrefix = "\\h3{";
      state.paragraphPostfix = "}";
      break;
    case "heading 4":
      state.paragraphPrefix = "\\h4{";
      state.paragraphPostfix = "}";
      break;
    case "heading 5":
      state.paragraphPrefix = "\\h5{";
      state.paragraphPostfix = "}";
      break;
    case "quote":
      state.paragraphPrefix = "\\begin{quote}\n";
      state.paragraphPostfix = "\n\\end{quote}";
      break;
    case "caption":
      state.paragraphPrefix = "\\caption{";
      state.paragraphPostfix = "}";
      break;
    case "list paragraph":
      let indentLevel, numId;
      const numPr = getXElement(paragraphProperties, "numPr");
      if (numPr !== undefined) {
        indentLevel = getXElement(numPr, "ilvl").getAttribute("w:val");
        numId = getXElement(numPr, "numId").getAttribute("w:val");
      }
      if (indentLevel !== undefined && numId !== undefined) {
        const listType = getListParagraphType(indentLevel, numId);
        state.paragraphPrefix = `\\item_${indentLevel}_${listType} `;
      }
      break;
    default:
      //それ以外のスタイルへの対応
      if (styleName.toLowerCase().includes("引用") || styleName.toLowerCase().includes("quot")) {
        state.paragraphPrefix = "\\begin{quote}\n";
        state.paragraphPostfix = "\n\\end{quote}";
      } else {
        //console.log("未処理の段落スタイル：" + styleName);
      }
      break;
  }
}

function processRun(runElement, state) {
  state.isMeaninglessRun = true;

  //画像
  const imageElement = runElement.querySelectorAll(":scope > drawing");
  for (let i = 0; i < imageElement.length; i++) {
    //画像それ自体にキャプション段落のスタイルが適用されている場合があるため、その対策
    if (state.paragraphPrefix === "\\caption{" && state.paragraphPostfix === "}") {
      state.paragraphPrefix = "";
      state.paragraphPostfix = "";
    }
    //画像それ自体に見出しのスタイルが適用されている場合があるため、その対策
    if (state.paragraphPrefix.startsWith("\\h")) {
      state.paragraphPostfix += processImage(imageElement[i]);
    } else {
      state.runPrefix += processImage(imageElement[i]);
    }
    state.isMeaninglessRun = false;
  }

  //行外画像のキャプション
  const mcfallbackElement = runElement.querySelectorAll(":scope > AlternateContent");
  for (let i = 0; i < mcfallbackElement.length; i++) {
    const txbxContents = mcfallbackElement[i].querySelectorAll("txbxContent"); //直下ではない
    if (txbxContents !== undefined && txbxContents.length > 0) {
      const mcParagraphs = txbxContents[0].querySelectorAll(":scope > p");
      for (let j = 0; j < mcParagraphs.length; j++) {
        state.paragraphPostfix += processParagraph(mcParagraphs[j]).trim(); // \caption{}が入る・・・captionじゃないときもある（学振のファイル等）
      }
    }
  }

  //文節ごとのフォント設定を取得
  const rPrItems = runElement.querySelectorAll(":scope > rPr");
  processRunProperties(rPrItems, state); //

  //フィールドコードの処理
  const fcItems = runElement.querySelectorAll(":scope > instrText");
  if (fcItems.length > 0) {
    fcItems.forEach(function (fcItem) {
      state.fcString += fcItem.textContent;
    });
  } else {
    if (state.fcString.trim() !== "") {
      processFieldCode(state); //
    }
    state.fcString = "";
  }

  //テキストを取得
  const textItems = runElement.querySelectorAll(":scope > t");
  for (let i = 0; i < textItems.length; i++) {
    //lengthは0か1
    //括弧の処理、エスケープ文字の処理（xml）、エスケープ文字の処理（TeX）、ダッシュの処理
    state.runPrefix += replaceBracket(
      replaceDashCharacters(escapeLaTeXCharacters(decodeXmlText(textItems[i].textContent)))
    );
    state.isMeaninglessRun = false;
  }

  //ルビ付きテキストの処理
  const rubyItems = runElement.querySelectorAll(":scope > ruby");
  for (let i = 0; i < rubyItems.length; i++) {
    const rubyText = getXElement(rubyItems[i], "rt");
    const rubyBase = getXElement(rubyItems[i], "rubyBase");
    //括弧の処理、エスケープ文字の処理（xml）、エスケープ文字の処理（TeX）、ダッシュの処理
    state.runPrefix +=
      "\\ruby{" +
      replaceBracket(replaceDashCharacters(escapeLaTeXCharacters(decodeXmlText(rubyBase.textContent)))) +
      "}{" +
      replaceBracket(replaceDashCharacters(escapeLaTeXCharacters(decodeXmlText(rubyText.textContent)))) +
      "}";
    state.isMeaninglessRun = false;
  }

  // 脚注参照・文末脚注参照の処理
  processNoteReference(runElement, state, "footnote", footnoteXml);
  processNoteReference(runElement, state, "endnote", endnoteXml);
}

// 脚注・文末脚注の共通の処理
function processNoteReference(runElement, state, noteType, noteXml) {
  const noteReferences = [...runElement.querySelectorAll(`:scope > ${noteType}Reference`)];

  for (const noteReference of noteReferences) {
    const noteId = noteReference.getAttribute("w:id");
    const noteItem = getXElementByAttribute(noteXml, noteType, "w:id", noteId);

    state.runPostfix += `\\${noteType}{`;

    const noteParagraphs = [...noteItem.querySelectorAll(":scope > p")];
    for (const noteParagraph of noteParagraphs) {
      const noteParagraphText = processParagraph(noteParagraph, true).trimStart();
      if (noteParagraphText !== "\\meaninglessparagraph") {
        state.runPostfix += noteParagraphText;
      }
    }

    if (state.runPostfix.endsWith(" \\\\\\indent ")) {
      state.runPostfix = state.runPostfix.slice(0, -11);
    }

    state.runPostfix += "\\nolinebreak}"; // 注で余分な改行が入ることを防ぐために追加
    state.isMeaninglessRun = false;
  }
}

// 文節ごとの装飾を処理
function processRunProperties(rPrItems, state) {
  for (const rPrItem of rPrItems) {
    processEastAsianLayout(rPrItem, state);
    processRunStyle(rPrItem, state);
  }
}

// 割り注と縦中横の処理
function processEastAsianLayout(rPrItem, state) {
  const eastAsianItems = [...rPrItem.querySelectorAll(":scope > eastAsianLayout")];
  for (const eastAsianItem of eastAsianItems) {
    const combine = eastAsianItem.getAttribute("w:combine");
    const vert = eastAsianItem.getAttribute("w:vert");
    const eastAsianID = eastAsianItem.getAttribute("w:id");

    let tag = null;
    if (combine === "1") tag = "\\warichu{";
    if (vert === "1" && isTate) tag = "\\tatechuyoko{";

    if (tag) {
      if (state.eastAsianID === eastAsianID) {
        state.paragraphPrefix = state.paragraphPrefix.slice(0, -1);
        state.runPostfix = "}" + state.runPostfix;
      } else {
        state.runPrefix += tag;
        state.runPostfix = "}" + state.runPostfix;
        state.eastAsianID = eastAsianID;
      }
    }
  }
}

// 文節の装飾の処理
function processRunStyle(rPrItem, state) {
  const styles = [
    {
      selector: ":scope > vertAlign",
      prefix: "\\textsubscript{",
      postfix: "\\textsubscript}",
      attribute: "w:val",
      include: ["subscript"],
    },
    {
      selector: ":scope > vertAlign",
      prefix: "\\textsuperscript{",
      postfix: "\\textsuperscript}",
      attribute: "w:val",
      include: ["superscript"],
    },
    { selector: ":scope > b", prefix: "\\textbf{", postfix: "\\textbf}", attribute: "w:val", exclude: "0" },
    { selector: ":scope > i", prefix: "\\textit{", postfix: "\\textit}", attribute: "w:val", exclude: "0" },
    /*{
      selector: ":scope > u",
      prefix: "%下線を使用する場合はudline.styを使うオプションを使用してください。\n",
      postfix: "",
      attribute: "w:val",
      exclude: "none",
    },
    {
      selector: ":scope > strike",
      prefix: "%取り消し線を使用する場合はudline.styを使うオプションを使用してください。\n",
      postfix: "",
      attribute: "w:val",
      exclude: "none",
    },*/
    { selector: ":scope > bdr", prefix: "\\fbox{", postfix: "\\fbox}" },
    {
      selector: ":scope > em",
      prefix: "\\kenten{",
      postfix: "\\kenten}",
      attribute: "w:val",
      include: ["dot", "comma"],
    },
  ];

  for (const style of styles) {
    const items = [...rPrItem.querySelectorAll(style.selector)];
    for (const item of items) {
      const attributeValue = item.getAttribute(style.attribute || "");
      if (
        (style.include && style.include.includes(attributeValue)) || // 圏点・上付き・下付き
        (style.exclude && // 太字・イタリック
          attributeValue !== style.exclude &&
          !state.paragraphPrefix.startsWith("\\h") && // 見出しには適用しない
          !state.paragraphPrefix.startsWith("\\title")) || // タイトルには適用しない
        !style.attribute // 囲み文字
      ) {
        state.runPrefix += style.prefix;
        state.runPostfix = style.postfix ? `${style.postfix}${state.runPostfix}` : state.runPostfix;
        break;
      }
    }
  }
}

// フィールドコードの処理
function processFieldCode(state) {
  //console.log(state.fcString);
  let processedText;
  const results_ruby = state.fcString.match(/hps[0-9]{1,2} \\o\\ad\(\\s\\up [0-9]{1,2}\((.*?)\),(.*?)\)/i);
  const results_cancel1 = state.fcString.match(/eq \\o\s?\((.*?),\/\)/i);
  const results_cancel2 = state.fcString.match(/eq \\o\s?\((.*?),／\)/i);
  const results_circled = state.fcString.match(/eq \\o\s?\\ac\(○,(.*?)\)/i);
  const results_overline = state.fcString.match(/eq \\x\s?\\to \((.*?)\)/i);
  const results_toc_level = state.fcString.match(/toc \\o "(.*?)-(.*?)"/i);
  const results_toc_any = state.fcString.match(/toc \\o/i);
  const results_toc_tables = state.fcString.match(/toc .* "figure"/i);
  const results_toc_figures = state.fcString.match(/toc .* "table"/i);

  if (results_ruby) {
    processedText = `\\ruby{${results_ruby[2]}}{${results_ruby[1]}}`;
  }
  if (results_cancel1) {
    processedText = `\\cancel{${results_cancel1[1]}}`;
    isCancelUsed = true;
  }
  if (results_cancel2) {
    processedText = `\\cancel{${results_cancel2[1]}}`;
    isCancelUsed = true;
  }
  if (results_circled) {
    processedText = `\\textcircled{${results_circled[1]}}`;
  }
  if (results_overline) {
    processedText = `$\\overline{${results_overline[1]}}$`;
  }
  if (results_toc_any) {
    if (results_toc_level) {
      processedText = `\\setcounter{tocdepth}{${results_toc_level[2]}}\n`;
    }
    processedText += "\\tableofcontents";
  }
  if (results_toc_tables) {
    processedText = "\\listoftables";
  }
  if (results_toc_figures) {
    processedText = "\\listoffigures";
  }

  if (processedText === undefined) {
    console.log("未処理のフィールドコード：" + state.fcString);
    return;
  } else if (processedText !== "") {
    //console.log(processedText);
  }

  if (
    processedText.includes("\\tableofcontents") ||
    processedText.includes("\\listoffigures") ||
    processedText.includes("\\listoftables")
  ) {
    state.paragraphPostfix += processedText;
    state.isMeaninglessParagraph = false;
  } else {
    state.runPrefix += processedText;
    state.isMeaninglessRun = false;
  }
}

// 数式を処理
function processMath(mathXml, isParagraphMath) {
  if (mathXml === null || mathXml === undefined) return;

  if (mathXml.nodeName === "m:oMathPara") {
    mathXml = mathXml.childNodes[0];
  } else if (mathXml.nodeName === "m:oMath") {
  } else return;

  isAmsMathUsed = true;
  let mathString = " ";

  if (isConvertMath) {
    mathString = processMathNode(mathXml);
    console.log("------");
  }

  if (isParagraphMath) {
    return "\\[" + mathString + "\\]";
  } else {
    return "\\(" + mathString + "\\)";
  }
}

// 画像ファイルを処理
function processImage(imageElement) {
  let width = 0,
    height = 0;
  isGraphicxUsed = true;

  //サイズの取得
  const imageExtents = imageElement.querySelectorAll("extent"); //直下ではない
  for (let i = 0; i < imageExtents.length; i++) {
    width = Math.round((parseFloat(imageExtents[i].getAttribute("cx")) / 914400) * 2.54 * 10) / 10; //20th of a Point	 / Inches / Centimeters
    height = Math.round((parseFloat(imageExtents[i].getAttribute("cy")) / 914400) * 2.54 * 10) / 10;
  }

  //ファイル名の取得
  const imageFileName = imageElement.querySelectorAll("blip"); //直下ではない
  for (let i = 0; i < imageFileName.length; i++) {
    const imageFileId = imageFileName[i].getAttribute("r:embed");
    const imageRelationshipName = getXElementByAttribute(relsXml, "Relationship", "Id", imageFileId);
    if (imageRelationshipName !== undefined) {
      const imageFileNameInDocx = imageRelationshipName.getAttribute("Target");
      if (imageFileNameInDocx !== null && imageFileNameInDocx.startsWith("media/")) {
        const imageFileNameInZip = "images/" + imageFileNameInDocx.slice(6).trim();
        return `\\includegraphics[width=${width}cm]{${imageFileNameInZip}}`;
      }
    }
  }
  return "Error: No image file found";
}

//リスト段落のナンバリング書式を読み込む
function getListParagraphType(indentLevel, numId, isRecursive = false) {
  const numElement = getXElementByAttribute(numberingXml, "num", "w:numId", numId);
  const abstractNumId = getXElement(numElement, "abstractNumId").getAttribute("w:val");
  const abstractNumElement = getXElementByAttribute(numberingXml, "abstractNum", "w:abstractNumId", abstractNumId);
  const levelElement = getXElementByAttribute(abstractNumElement, "lvl", "w:ilvl", indentLevel);

  //↑が失敗する場合の対処
  if (levelElement === undefined) {
    const styleID = getXElement(abstractNumElement, "numStyleLink").getAttribute("w:val");
    const styleElement = getXElementByAttribute(styleXml, "style", "w:styleId", styleID);
    const linkedNumId = styleElement.querySelectorAll("numId")[0].getAttribute("w:val"); //queryselectorが失敗した場合の処理？
    if (!isRecursive) {
      return getListParagraphType(indentLevel, linkedNumId, true);
    } else {
      console.log("2x recursive call in getListParagraphType:");
      return 0;
    }
  }

  const types = getXElement(levelElement, "numFmt").getAttribute("w:val");

  if (types === undefined) return 0; //不明
  else if (types === "bullet") return 0; //丸印
  else return 1; //数字
}

// XMLから直下の要素を1つだけ取得
function getXElement(xmlElement, tagName) {
  //console.log(xmlElement);
  const elem = xmlElement.querySelectorAll(":scope > " + tagName);

  if (elem.length === 1) {
    return elem[0];
  } else if (elem.length === 0) {
    //console.log('getXElement: no ' + tagName + ' element in ' + xmlElement.nodeName);
    return;
  } else {
    console.log("getXElement: more than 1 " + tagName + " elements in " + xmlElement);
    return;
  }
}

// XMLから直下の要素を（属性を指定して）1つだけ取得
function getXElementByAttribute(xmlElement, tagName, attributeName, attributeText) {
  //console.log(xmlElement);
  const elem = Array.from(xmlElement.querySelectorAll(":scope > " + tagName)).filter(
    (el) => el.getAttribute(attributeName) === attributeText
  );

  if (elem.length === 1) {
    return elem[0];
  } else if (elem.length === 0) {
    //console.log('getXElementByAttribute: no ' + tagName + ' element in ' + xmlElement.nodeName);
    return;
  } else {
    console.log("getXElementByAttribute: more than 1 " + tagName + " elements in " + xmlElement);
    return;
  }
}

// XMLをパースして、ルート要素を返す
async function getXmlRoot(xmlFile) {
  if (xmlFile === undefined) return undefined;

  try {
    xmlText = await xmlFile.async("string");
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, "text/xml");
    return doc.documentElement;
  } catch (error) {
    console.log("ERROR in getXmlRoot:");
    console.log(error);
  }
}

// LaTeXの特殊文字をエスケープ
function escapeLaTeXCharacters(text) {
  text = text.replace(/\\/g, "\\textbackslash ");
  text = text.replace(/#/g, "\\#");
  text = text.replace(/\$/g, "\\$");
  text = text.replace(/%/g, "\\%");
  text = text.replace(/&/g, "\\&");
  text = text.replace(/~/g, "\\textasciitilde ");
  text = text.replace(/_/g, "\\_");
  text = text.replace(/\^/g, "\\textasciicircum ");
  text = text.replace(/{/g, "\\{");
  text = text.replace(/}/g, "\\}");
  text = text.replace(/\|/g, "\\textbar ");
  text = text.replace(/</g, "\\textless ");
  text = text.replace(/>/g, "\\textgreater ");
  return text;
}

// ダーシの文字を揃える
function replaceDashCharacters(text) {
  text = text.replace(/──/g, "――");
  text = text.replace(/——/g, "――");
  return text;
}

// 括弧を変換
function replaceBracket(text) {
  //if (bFullWidthBracket) {
  text = text.replace(/\(/g, "（");
  text = text.replace(/\)/g, "）");
  //}
  return text;
}

// XMLをデコード
function decodeXmlText(text) {
  //既にデコードされた形になっているので、そのまま引数を返す
  //デコードが必要なら、下記
  //let span = document.createElement("SPAN");
  //span.innerHTML = text;
  //return span.innerText;
  return text;
}

// 括弧内の文字列を取得
// str内から、beforestrで始まる括弧の中の文字列を取得
// 括弧の記号はopenchr、closechrで指定する
function inParenString(str, beforestr, openchr, closechr) {
  let inParenString = "";
  let openParenCount = 0;
  let closeParenCount = 0;

  let n = str.indexOf(beforestr + openchr);
  if (n < 0) return ""; // beforestr + openchrがない場合は""を返す

  openParenCount++;

  for (let i = n + beforestr.length + 1; i < str.length; i++) {
    let c = str.substring(i, i + 1);
    if (c === openchr.toString()) {
      openParenCount++;
    } else if (c === closechr.toString()) {
      closeParenCount++;
    }
    if (openParenCount === closeParenCount) return inParenString;
    inParenString += c;
  }

  return "";
}

// 文字列のなかに指定した文字（列）が登場する回数を数える
function countChar(str, char) {
  return str.split(char).length - 1;
}

// 注以外の箇所の括弧・亀甲括弧内をrelsizeで囲む
function reduceFontSizeWithinParentheses(str) {
  let textarray1 = splitStringByTag(str, "\\footnote");
  let textarray2 = splitStringArrayByTag(textarray1, "\\endnote");
  let textarray3 = splitStringArrayByTag(textarray2, "\\caption");

  for (let i = 0; i < textarray3.length; i++) {
    if (textarray3[i].indexOf("note{") === -1 && textarray3[i].indexOf("\\caption{") === -1) {
      textarray3[i] = textarray3[i]
        .replace(/（/g, "\\relsize{-0.5}（")
        .replace(/）/g, "）\\relsize{0.5}")
        .replace(/〔/g, "\\relsize{-0.5}〔")
        .replace(/〕/g, "〕\\relsize{0.5}");
    }
  }
  return textarray3.join("");
}

// 指定された文字列をタグで区切って配列に格納
function splitStringByTag(str, tag) {
  let textarray = new Array();
  textarray.push(str);
  return splitStringArrayByTag(textarray, tag);
}

// 指定された文字列の配列をさらにタグで区切って配列に格納
function splitStringArrayByTag(textArray, tag) {
  let textArray2 = new Array();
  for (let i = 0; i < textArray.length; i++) {
    if (textArray[i].includes(tag)) {
      while (textArray[i].includes(tag)) {
        const tagstr = tag + "{" + inParenString(textArray[i], tag, "{", "}") + "}";
        const tagstrindex = textArray[i].indexOf(tagstr);
        const text1 = textArray[i].substr(0, tagstrindex);
        const text2 = textArray[i].substr(tagstrindex, tagstr.length);
        const text3 = textArray[i].substr(tagstrindex + tagstr.length);
        if (text1 !== "") textArray2.push(text1);
        if (text2 !== "") textArray2.push(text2);
        textArray[i] = text3;
      }
      textArray2.push(textArray[i]);
    } else {
      textArray2.push(textArray[i]);
    }
  }
  return textArray2;
}

// 図や表のキャプションの不要な部分（図１など）を削除する
function RemoveCaptionHeader(strCaption) {
  //console.log("before: " + strCaption);
  const delimiterRegex = new RegExp(["　", " "].join("|"), "g");
  const strArray = strCaption.split(delimiterRegex);
  const headerStrs = ["Table", "Figure", "図", "表"];
  const strArrayLength = strArray.length;

  for (const headerStr of headerStrs) {
    if (strArray[0] !== null && strArray[0].trim() === headerStr) {
      strArray[0] = "";
      if (strArrayLength > 1 && strArray[1] !== null && /^\d+$/.test(strArray[1])) {
        strArray[1] = "";
      }
      break;
    }
  }
  //console.log("after:  " + strArray.join(" ").trim());
  return strArray.join(" ").trim();
}

// テーブルのXMLを処理
function processTable(tableElement) {
  let columnNumber = 0,
    rowNumber = 0,
    columnMaxNumber = 0,
    footnoteCount = 0,
    endnoteCount = 0;
  let tablePrefix = "\\begin{tabular}";
  let tablePostfix = "";
  let tableText = "";
  let footnotePostfix = "",
    endnotePostfix = "";

  const rowElements = tableElement.querySelectorAll(":scope > tr");
  for (let i = 0; i < rowElements.length; i++) {
    rowNumber++;
    columnNumber = 0;
    let rowPrefix = "",
      rowPostfix = "";

    const columnElements = rowElements[i].querySelectorAll(":scope > tc");
    for (let j = 0; j < columnElements.length; j++) {
      console.log(i + " : " + j); //////////////////

      let isMergedEmptyCell = false;
      let mergedRowsCount = 0;
      let mergedColumnsCount = 0;
      let columnPrefix = "",
        columnPostfix = "";

      const gridSpanElements = columnElements[j].querySelectorAll("gridSpan"); // 直下ではない
      let gridSpanText = null;

      if (gridSpanElements.length > 0) {
        gridSpanText = gridSpanElements[0].getAttribute("w:val");
      }
      if (gridSpanText !== null) {
        mergedColumnsCount = parseInt(gridSpanText);
        columnNumber += mergedColumnsCount;
        isMultiRowUsed = true;
      } else {
        columnNumber++;
      }

      const vmergeElements = columnElements[j].querySelectorAll("vMerge"); // 直下ではない
      if (vmergeElements.length > 0) {
        const vmergeElement_restart = vmergeElements[0].getAttribute("w:val");
        if (vmergeElement_restart !== null) {
          mergedRowsCount = getMergedRowsLength(tableElement, columnNumber, rowNumber);
        } else {
          isMergedEmptyCell = true;
        }
        isMultiRowUsed = true;
      }

      if (isMergedEmptyCell) {
        console.log("isMergedEmptyCell = true"); ///////////
        let k = mergedColumnsCount;
        console.log("mrowcount: " + mergedRowsCount); ///////
        console.log("mcolcount: " + mergedColumnsCount); ///////
        if (rowPrefix === "") k--;
        console.log("k: " + k); ///////
        for (let l = 1; l <= k; l++) columnPrefix += " &";
        if (columnPrefix === "" && j > 0) columnPrefix = " &";
        console.log("columnPrefix : " + columnPrefix); ///////////
      } else {
        if (rowPrefix !== "") columnPrefix += " & ";

        if (mergedRowsCount > 0) {
          if (mergedColumnsCount > 0) {
            columnPrefix += "\\multicolumn{" + mergedColumnsCount + "}{c}{\\multirow{" + mergedRowsCount + "}{*}{";
            columnPostfix = "}}";
          } else {
            columnPrefix += "\\multirow{" + mergedRowsCount + "}{*}{";
            columnPostfix = "}";
          }
        } else {
          if (mergedColumnsCount > 0) {
            columnPrefix += "\\multicolumn{" + mergedColumnsCount + "}{c}{";
            columnPostfix = "}";
          } else {
            columnPrefix += "";
          }
        }
        console.log("columnPrefix : " + columnPrefix); ///////////
      }

      //セル内paragraph読み込み
      const paragraphsInCell = columnElements[j].querySelectorAll(":scope > p");

      if (paragraphsInCell !== undefined) {
        //セル内改行はtabularのネストで対応
        if (paragraphsInCell.length > 1) {
          columnPrefix += "\\begin{tabular}{c}";
          columnPostfix = "\\end{tabular}" + columnPostfix;
        }

        for (let m = 0; m < paragraphsInCell.length; m++) {
          //セル内パラグラフを取得
          let paragraphTextInCell = processParagraph(paragraphsInCell[m], false, true).trim();

          if (paragraphTextInCell !== "\\meaninglessparagraph" && paragraphTextInCell.length > 0) {
            //セル内改行はtabularのネストで対応
            if (paragraphsInCell.length > 1) paragraphTextInCell += " \\\\ ";

            // テーブル内は強制的に行内数式にする（\[ \]→\( \)の置換処理）
            paragraphTextInCell = paragraphTextInCell.replace(/\\\[/g, "\\(");
            paragraphTextInCell = paragraphTextInCell.replace(/\\\]/g, "\\)");

            let noteText = inParenString(paragraphTextInCell, "note", "{", "}");
            while (noteText !== "") {
              let footorend = paragraphTextInCell.substr(paragraphTextInCell.indexOf("note{" + noteText) - 1, 1);
              if (footorend === "t") {
                footnotePostfix += "\\addtocounter{footnote}{1}\\footnotetext{" + noteText + "}\n";
                paragraphTextInCell =
                  paragraphTextInCell.substr(0, paragraphTextInCell.indexOf("\\footnote{" + noteText) + 9) +
                  "mark " +
                  paragraphTextInCell.substr(
                    paragraphTextInCell.indexOf("\\footnote{" + noteText) + noteText.length + 11
                  );
                footnoteCount++;
              } else if (footorend === "d") {
                endnotePostfix += "\\addtocounter{endnote}{1}\\endnotetext{" + noteText + "}\n";
                paragraphTextInCell =
                  paragraphTextInCell.substr(0, paragraphTextInCell.indexOf("\\endnote{" + noteText) + 8) +
                  "mark " +
                  paragraphTextInCell.substr(
                    paragraphTextInCell.indexOf("\\endnote{" + noteText) + noteText.length + 10
                  );
                endnoteCount++;
              } else console.log("処理できないノート: " + footorend);
              noteText = inParenString(paragraphTextInCell, "note", "{", "}");
            }
          } else {
            paragraphTextInCell = " ";
          }

          columnPrefix += paragraphTextInCell;
          console.log(paragraphTextInCell); /////////////////////
        }
      }
      rowPrefix += columnPrefix + columnPostfix;
    }

    if (columnMaxNumber < columnNumber) columnMaxNumber = columnNumber;

    tableText += rowPrefix + rowPostfix + " \\\\\n";
  }

  const repeatedString = "l".repeat(columnMaxNumber);
  tablePrefix = tablePrefix + "{" + repeatedString + "}\n\\hline";

  tablePostfix = "\\hline\n\\end{tabular}";

  if (footnoteCount > 0) {
    tablePostfix += "\n\\addtocounter{footnote}{-" + footnoteCount + "}\n" + footnotePostfix;
  }
  if (endnoteCount > 0) {
    tablePostfix += "\n\\addtocounter{endnote}{-" + endnoteCount + "}\n" + endnotePostfix;
  }

  return tablePrefix + "\n" + tableText + tablePostfix;
}

// テーブル内で縦方向に結合されたセルの数を調べる（最適化後）
function getMergedRowsLength(tableElement, mergedColumnStart, mergedRowStart) {
  let mergedRowsCount = 1;
  const rows = tableElement.querySelectorAll(":scope > tr");

  for (let i = mergedRowStart; i < rows.length; i++) {
    const row = rows[i];
    const cells = row.querySelectorAll(":scope > tc");

    for (let j = 0; j < cells.length; j++) {
      const cell = cells[j];
      const vMerge = cell.querySelector("vMerge");

      if (j === mergedColumnStart - 1) {
        if (vMerge && !vMerge.getAttribute("w:val")) {
          mergedRowsCount++;
        } else {
          return mergedRowsCount;
        }
      }
    }
  }

  return mergedRowsCount;
}

//
//最適化前の関数を保存
//
/*
// テーブル内で縦方向に結合されたセルの数を調べる
function getMergedRowsLength_old(tableElement, mergedColumnStart, mergedRowStart) {
  let mergedRowsCount = 1;
  let columnNumber = 0,
    rowNumber = 0;
  let isEnd = false;

  const rowsElement = tableElement.querySelectorAll(":scope > tr");
  for (let i = 0; i < rowsElement.length; i++) {
    rowNumber++;
    columnNumber = 0;

    if (rowNumber > mergedRowStart) {
      const columnsElement = rowsElement[i].querySelectorAll(":scope > tc");

      for (let j = 0; j < columnsElement.length; j++) {
        const gridSpanElement = columnsElement[j].querySelectorAll("gridSpan"); // 直下ではない
        if (gridSpanElement.length > 0) {
          const gridSpanAttribute = gridSpanElement[0].getAttribute("w:val");
          if (gridSpanAttribute !== null) {
            columnNumber += parseInt(gridSpanAttribute);
          }
        } else {
          columnNumber++;
        }
        if (columnNumber === mergedColumnStart) {
          const vmergeElement = columnsElement[j].querySelectorAll("vMerge"); // 直下ではない
          if (vmergeElement.length > 0) {
            const vmergeElementAttribute = vmergeElement[0].getAttribute("w:val");
            if (vmergeElementAttribute === null) {
              if (!isEnd) mergedRowsCount++;
            } else {
              isEnd = true;
            }
          }
        }
      }
    }
  }
  return mergedRowsCount;
}

function processRunProperties_old(rPrItems, state) {
  for (let i = 0; i < rPrItems.length; i++) {
    //割注<w:eastAsianLayout w:id='-1182637056' w:combine='1'/>　縦中横<w:eastAsianLayout w:id='2' w:vert='on' />
    const eastAsianItems = rPrItems[i].querySelectorAll(":scope > eastAsianLayout");
    for (let j = 0; j < eastAsianItems.length; j++) {
      switch (eastAsianItems[j].getAttribute("w:combine")) {
        case "1":
          if (state.eastAsianID === eastAsianItems[j].getAttribute("w:id")) {
            state.paragraphPrefix = state.paragraphPrefix.slice(0, -1); //[..^1];
            state.runPostfix = "}" + state.runPostfix;
          } else {
            state.runPrefix += "\\warichu{";
            state.runPostfix = "}" + state.runPostfix;
            state.eastAsianID = eastAsianItems[j].getAttribute("w:id");
          }
          break;
        default:
          break;
      }

      switch (eastAsianItems[j].getAttribute("w:vert")) {
        case "1":
          if (isTate) break; // 縦書きでない場合は\tatechuyokoは無効
          if (state.eastAsianID === eastAsianItems[j].getAttribute("w:id")) {
            state.paragraphPrefix = state.paragraphPrefix.slice(0, -1); //[..^1];
            state.runPostfix = "}" + state.runPostfix;
          } else {
            state.runPrefix += "\\tatechuyoko{";
            state.runPostfix = "}" + state.runPostfix;
            state.eastAsianID = eastAsianItems[j].getAttribute("w:id");
          }
          break;
        default:
          break;
      }
    }

    //上付き下付きw:vertAlign w:val='subscript'/'superscript'
    const vertAlignItems = rPrItems[i].querySelectorAll(":scope > vertAlign");
    for (let j = 0; j < vertAlignItems.length; j++) {
      switch (vertAlignItems[j].getAttribute("w:val")) {
        case "subscript":
          //後で重複を削除するためにpostfixにもタグを入れておく（paragraphTextFinalizedを見よ。以下同様）
          state.runPrefix += "\\textsubscript{";
          state.runPostfix = "\\textsubscript}" + state.runPostfix;
          break;
        case "superscript":
          state.runPrefix += "\\textsuperscript{";
          state.runPostfix = "\\textsuperscript}" + state.runPostfix;
          break;
        default:
          //Console.Write(vertAlign.Attribute(w + 'val').Value);
          break;
      }
    }

    //太字w:b
    const boldItems = rPrItems[i].querySelectorAll(":scope > b");
    for (let j = 0; j < boldItems.length; j++) {
      //w:val='0'を除外
      if (boldItems[j].getAttribute("w:val") === "0") break;
      state.runPrefix += "\\textgt{";
      state.runPostfix = "\\textgt}" + state.runPostfix;
      break;
    }

    //イタリックw:i
    const italicItems = rPrItems[i].querySelectorAll(":scope > i");
    for (let j = 0; j < italicItems.length; j++) {
      //w:val='0'を除外
      if (italicItems[j].getAttribute("w:val") === "0") break;
      state.runPrefix += "\\textit{";
      state.runPostfix = "\\textit}" + state.runPostfix;
      break;
    }

    //下線：w:u
    const underlineItems = rPrItems[i].querySelectorAll(":scope > u");
    for (let j = 0; j < underlineItems.length; j++) {
      //w:val='none'を除外
      if (underlineItems[j].getAttribute("w:val") === "none") break;
      if (false) {
        state.runPrefix += "\\ul{";
        state.runPostfix = "\\ul}" + state.runPostfix;
      } else {
        state.runPrefix = "%下線を使用する場合はudline.styを使うオプションを使用してください。\n" + state.runPrefix;
      }
      break;
    }

    //取り消し線<w:strike w:val='0'/>
    const strikelineItems = rPrItems[i].querySelectorAll(":scope > strike");
    for (let j = 0; j < strikelineItems.length; j++) {
      //w:val='none'を除外
      if (strikelineItems[j].getAttribute("w:val") === "none") break;
      if (false) {
        // if bUdline＊＊
        if (!state.runPrefix.includes("\\ul{")) {
          // udline.styの\ul{} \Sl{}は同時使用不可のため
          state.runPrefix += "\\Sl{";
          state.runPostfix = "\\Sl}" + state.runPostfix;
        }
      } else {
        state.runPrefix =
          "%取り消し線を使用する場合はudline.styを使うオプションを使用してください。\n" + state.runPrefix;
      }
      break;
    }

    //ボックスw:bdr
    const bdrItems = rPrItems[i].querySelectorAll(":scope > bdr");
    for (let j = 0; j < bdrItems.length; j++) {
      state.runPrefix += "\\fbox{";
      state.runPostfix = "\\fbox}" + state.runPostfix;
      break;
    }

    //強調（圏点、luatexja-rubyの仕様上、最後に処理する必要がある）
    const emItems = rPrItems[i].querySelectorAll(":scope > em");
    for (let j = 0; j < emItems.length; j++) {
      switch (emItems[j].getAttribute("w:val")) {
        case "dot":
        case "comma":
          state.runPrefix += "\\kenten{";
          state.runPostfix = "\\kenten}" + state.runPostfix;
          break;
        default:
          //Console.Write(em.Attribute(w + 'val').Value);
          break;
      }
    }
  }
}

function old_foot_endnote() {
  //脚注参照
  const footnotesItem = runElement.querySelectorAll(":scope > footnoteReference");
  for (let i = 0; i < footnotesItem.length; i++) {
    const noteItem = getXElementByAttribute(footnoteXml, "footnote", "w:id", footnotesItem[i].getAttribute("w:id"));
    state.runPostfix += "\\footnote{";
    const footnoteParagraphs = noteItem.querySelectorAll(":scope > p");
    for (let j = 0; j < footnoteParagraphs.length; j++) {
      const noteParagraphText = processParagraph(footnoteParagraphs[j], true).trimStart();
      if (noteParagraphText !== "\\meaninglessparagraph") {
        state.runPostfix = state.runPostfix + noteParagraphText;
      }
    }
    if (state.runPostfix.endsWith(" \\\\\\indent ")) state.runPostfix = state.runPostfix.slice(0, -11);

    state.runPostfix += "\\nolinebreak}"; //注で余分な改行が入ることを防ぐために追加
    state.isMeaninglessRun = false;
  }

  //文末脚注参照
  const endnotesItem = runElement.querySelectorAll(":scope > endnoteReference");
  for (let i = 0; i < endnotesItem.length; i++) {
    const noteItem = getXElementByAttribute(endnoteXml, "endnote", "w:id", endnotesItem[i].getAttribute("w:id"));
    state.runPostfix += "\\endnote{";
    const endnoteParagraphs = noteItem.querySelectorAll(":scope > p");
    for (let j = 0; j < endnoteParagraphs.length; j++) {
      const noteParagraphText = processParagraph(endnoteParagraphs[j], true).trimStart();
      if (noteParagraphText !== "\\meaninglessparagraph") {
        state.runPostfix = state.runPostfix + noteParagraphText;
      }
    }
    if (state.runPostfix.endsWith(" \\\\\\indent ")) state.runPostfix = state.runPostfix.slice(0, -11);

    state.runPostfix += "\\nolinebreak}"; //注で余分な改行が入ることを防ぐために追加
    state.isMeaninglessRun = false;
  }
}
*/
