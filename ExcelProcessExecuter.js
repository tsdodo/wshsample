/**
 * 開いたexcelファイルで実行する処理が記載されたfunctionを渡します。
 * 
 * @constructor
 * @classdesc 指定されたフォルダのExcelを開いて処理を実行します。
 * @param {function(ExcelProcessExecuter,WorkBooks)} preProcessSheet シート処理前の処理関数
 * @param {function(ExcelProcessExecuter,WorkBooks,WorkSheets)} processSheet シート毎の処理関数
 * @param {function(ExcelProcessExecuter,WorkBooks)} postProcessSheet シート処理後の処理関数
 * @param {boolean} saveFlg 開いたファイルを保存するかどうかのフラグ
 */
function ExcelProcessExecuter(preProcessSheet,processSheet,postProcessSheet,saveFlg){

    /**
     * シート処理前の処理関数
     * @type {function(ExcelProcessExecuter,WorkBooks)} 
     */
    this.preProcessSheet = preProcessSheet;

    /**
     * シート毎の処理関数
     * @type {function(ExcelProcessExecuter,WorkBooks,WorkSheets)} 
     */
    this.processSheet = processSheet;

    /**
     * シート処理後の処理関数
     * @type {function(ExcelProcessExecuter,WorkBooks)} 
     */
    this.postProcessSheet = postProcessSheet;

    /** 
     * 開いたファイルを保存するかどうかのフラグ
     * @type {boolean}
     */
    this.saveFlg = typeof saveFlg==='undefined'?true:saveFlg;

    /**
     * エクセルオブジェクト
     * @type {Excel.application}
     */
    this.xlso = WScript.CreateObject("Excel.Application");

    /**
     * ファイルシステムオブジェクト
     * @type {Scripting.FileSystemObject}
     */
    this.fso = WScript.CreateObject("Scripting.FileSystemObject");    

    /** 
     * フォルダの再帰処理をするかどうかのフラグ
     * @type {boolean}
     */
    this.recursiveFlg = true;

    /**
     * 正常処理されたexcelファイルの件数
     * @type {number}
     */
    this.nCnt = 0;

    /**
     * エラーになったexcelファイルの件数
     * @type {number}
     */
    this.eCnt = 0;

    /**
     * エラーメッセージ
     * @type {string}
     */
    this.errMsg = "";
}

/**
 * 指定されたフォルダ名のフォルダ内のexcelファイルを処理します。
 * 
 * @memberof ExcelProcessExecuter
 * @param {string} folderName　フォルダ名
 * @param {boolean} recursiveFlg　再帰処理フラグ（指定しない場合はtrue)
 */
ExcelProcessExecuter.prototype.doProcess = function(folderName,recursiveFlg) {
    this.doProcessByFSO(this.fso.GetFolder(folderName),recursiveFlg);
}

/**
 * 指定されたフォルダ名のフォルダ内のexcelファイルを処理します。
 * 
 * @memberof ExcelProcessExecuter
 * @param {object} folder　フォルダのFileSystemObject
 * @param {boolean} recursiveFlg　再帰処理フラグ（指定しない場合はtrue)
 */
ExcelProcessExecuter.prototype.doProcessByFSO = function(folder,recursiveFlg) { 
    //jscriptだとデフォルト引数使えないのでこうする・・
    this.recursiveFlg = typeof recursiveFlg === 'undefined'?true:recursiveFlg;

    var self = this;
    //フォルダ内のエクセルファイルに対して実行
    forEach (folder.Files,
        function(file) {
            var fileName = file.path;
            var extention = self.fso.GetExtensionName(fileName);
            if (extention == "xls" || extention == "xlsx"  || extention == "xlsm") {
                try {
                    var book = self.xlso.WorkBooks.Open(fileName);
                    //シート処理前処理
                    self.preProcessSheet(self,book);
                    //シートごとの処理
                    forEach (book.WorkSheets,function(sheet){self.processSheet(self,book,sheet)})
                    //シート処理後の処理
                    self.postProcessSheet(self,book);

                    if(self.saveFlg) {
                        book.save();
                    }
                } catch (e) {
                    self.errMsg = self.errMsg + "\n" + fileName + ":" + e.message;
                    self.eCnt++;
                    return; //普通のループならcontinueだけどfunctionにしたのでreturn
                } finally {
                    if (book != null) {
                        book.close();
                    }
                }
                self.nCnt++;
            }
        });
    //サブフォルダがあったら再帰処理
    if (self.recursiveFlg){
        forEach (folder.SubFolders,
            function(subFolder) {
                self.doProcessByFSO(subFolder);
            });    
    }
}

/**
 * ダイアログで選択されたフォルダ内のexcelファイルを処理します。
 * 
 * @memberof ExcelProcessExecuter
 * @param {boolean} 再帰処理フラグ（指定しない場合はtrue)
 */
ExcelProcessExecuter.prototype.doProcessByDialog = function(recursiveFlg) { 
    var sho = WScript.CreateObject("Shell.Application");
    var shoFolder = sho.BrowseForFolder(0, "処理対象のファイルがあるフォルダを選択してください。", 0);
    if (shoFolder != null ) {
        //実行(引数はフォルダ名)
        this.doProcess(shoFolder.items().item().path,recursiveFlg);
    }else{
        //キャンセルされたら終了
        this.errMsg = "\nフォルダの選択がキャンセルされました。";
    }
}


/**
 * excelを終了します。
 * 
 * @memberof ExcelProcessExecuter
 */
ExcelProcessExecuter.prototype.quit = function() {
    if(this.xlso != null){
        this.xlso.quit();
    }
}

/**
 * コレクションのアイテムについてループ処理を実行します。
 * 
 * @param {collection} enumarable コレクション
 * @param {function(object)} delegate コレクションの要素毎の処理関数
 */
function forEach(enumarable, delegate){
    var e = new Enumerator(enumarable);
    e.moveFirst();
    while(!e.atEnd()) {
        delegate(e.item());
        e.moveNext();
    }
}
