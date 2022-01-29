/*
rubyDeepLearning.jsx
(c)2017 Shock tm
ワードルビ文章をペーストした選択テキストフレームにルビを振る
それをxmlに蓄積していき深層学習していく
通常起動でルビデータベース修正
2017-01-05 ver1.0
2019-09-07 ver1.1 ファイル選択ルビ振りモード追加・初回のみルビ振りモード追加
2019-09-10 ver1.2 バグフィクス
*/
var xd;
var deepIniPath=decodeURI(Folder.temp+"/rubyDeep.xml");
var ini=ReadXML(deepIniPath)
var deepXmlPath=decodeURI(Folder.temp+"/rubyDeep.xml");
var rf=true;
rubyMain();
//取り消し用
function rubyMain(){
    app.doScript("doRubyMain()", ScriptLanguage.JAVASCRIPT, [], UndoModes.fastEntireScript);
}
//メーンルーチン
function doRubyMain(){
    try{
        rubyPath()
        if (app.documents.length != 0){
            var moto=ReadXML(deepXmlPath)
            if (app.selection.length!=0){
                for(i=0;i<app.selection.length;i++){
                    if(app.selection[i] == "[object TextFrame]"){
                        var tag=app.selection[i].parentStory
                        app.findGrepPreferences=NothingEnum.nothing;
                        app.findGrepPreferences.properties={findWhat:"~K+\\([ぁ-ゖ]*?\\)"};
                        var wds=tag.findGrep();
                        app.findGrepPreferences.properties={findWhat:"~K+\\([ぁ-ゖ]*?\\)([ぁ-ゖ]*)"};
                        var wdsPlus=tag.findGrep();
                        xd=wordRubySet(wds,wdsPlus)
                        if(rf){
                            app.findGrepPreferences.properties={findWhat:"~K+([ぁ-ゖ]*)"};
                            var only=tag.findGrep();
                            rubiesSet(only)
                        }
                    }
                }
                if(xd!=moto){
                    WriteText(deepXmlPath,xd,"UTF8")
                }
            }else{
                rubiDia("change")
            }
        }else{
            rubiDia("change")
        }
    }catch(e){
        alert("error:"+e,"rubyDeepLearning.jsx")
    }
}
//ルビ登録ダイアログ mode＝deppは学習 changeは新規修正 newはワード新規登録
function rubiDia(mode,rubies,obj){
    var newWindow= new Window ('dialog', "ルビ登録  writen by Shock tm", [  0,  0,  0+260,  0+275])
    var flag=false;var ind=0
    var selectMode=ini["selectmode"];
    var oneMode=ini["onemode"];
    with(newWindow){;center();
        var findStr=add('edittext',[ 70, 10, 70+110, 10+ 20],'',{multiline:false});
        var findBt=add('button',   [190, 10,190+ 70, 10+ 22],'検索');
        var sta1=add('statictext', [ 10, 10, 10+240, 10+ 14],'単語検索：',{multiline:true});
        var sta2=add('statictext', [ 10, 48, 10+ 60, 48+ 14],'親文字：',{multiline:false});
        var sta3=add('statictext', [ 70, 48, 70+ 60, 48+ 14],'' , {multiline:false});
        var sta4=add('statictext', [125, 48,125+ 40, 48+ 14],'候補：',{multiline:false});sta4.visible=false;
        var dla=add('dropdownlist',[160, 45,160+100, 45+ 20]);dla.visible=false;
        var dlo=add('dropdownlist',[  0,  0,  0+  0,  0+  0]);dlo.visible=false;
        var dlr=add('dropdownlist',[  0,  0,  0+  0,  0+  0]);dlr.visible=false;
        var sta5=add('statictext', [ 10, 78, 10+ 60, 75+ 14],'ルビ文字：',{multiline:false});
        var rubyStr=add('edittext',[ 70, 75, 70+185, 75+ 20],'',{multiline:false});
        var sta6=add('statictext', [ 10,108, 10+ 60,108+ 14],'送り仮名：',{multiline:false});
        var okuStr=add('edittext', [ 70,105, 70+185,105+ 20],'',{multiline:false});
        var quitBt=add('button',   [ 70,135, 70+ 50,135+ 25],'終了');
        var delBt=add('button',    [140,135,140+ 50,135+ 25],'削除');
        var chengeBt=add('button', [200,135,200+ 50,135+ 25],'登録',{name: 'ok'});
        var cansel=add('button',   [100,170,100+150,170+ 25],'Esc（登録せずにルビ振り）');cansel.visible=false
        var openBt=add('button',   [  5,170, 5+125,170+ 25],'登録ファイルを開く');
        var xmlBt=add('button',    [135,170,135+125,170+ 25],'登録ファイルの変更');
        var oneModeBt=add('button',[ 5,210,  5+255,210+ 25],'初回のみルビ振りモードに変更');
        var selModeBt=add('button',[ 5,240,  5+255,240+ 25],'ファイル選択ルビ振りモードに変更');
        if((mode=="new")|(mode=="deep")){
            findStr.visible=false
            openBt.visible=false
            xmlBt.visible=false
            findBt.visible=false
            delBt.visible=false
            quitBt.visible=false
            cansel.visible=true
            if(mode=="new"){
                sta1.text="テキスト「"+rubies[0]+"」のルビに区切り部分（欧文スペース）をどうぞ。（※「Esc」キーで登録せずにルビを振ります。）"
            }else{
                sta1.text="テキスト「"+rubies[0]+"」のルビ「"+rubies[1]+"」を送り仮名の調整をし、登録しますか？\n（※「Esc」キーで登録キャンセル。）"
                cansel.text="Esc（登録キャンセル）"
            }
            sta3.text=rubies[0]
            rubyStr.text=rubies[1]
            okuStr.text=rubies[2]
            rubyStr.active=true
        }else{
            findStr.active=true
        }
        if(selectMode=="true"){
            selModeBt.text="通常ルビ振りモードに戻す";
        }
        if(oneMode=="true"){
            oneModeBt.text="全回ルビ振りモードに戻す";
        }
        findBt.onClick=function(){
            dla.removeAll()
            dlo.removeAll()
            dlr.removeAll()
            dla.visible=false
            sta4.visible=false
            if(findStr.text!=""){
                sta3.text=findStr.text
                var tempRuby=xd[findStr.text]
                if(tempRuby.length()==0){
                    rubyStr.text=""
                    okuStr.text=""
                    ind=0
                    rubyStr.active=true
                }else{
                    sta4.visible=true
                    dla.visible=true
                    dlo.active=true
                    dla.add("item","新規...")
                    dlo.add("item","")
                    dlr.add("item","")
                    for(var se=0;se<tempRuby.length();se++){
                        if(tempRuby[se].@送り仮名.toString()!=""){
                            dla.add("item",tempRuby[se].toString()+" 【"+tempRuby[se].@送り仮名.toString()+"】")
                        }else{
                            dla.add("item",tempRuby[se].toString())
                        }
                        dlo.add("item",tempRuby[se].@送り仮名.toString())
                        dlr.add("item",tempRuby[se].toString())
                    }
                    dla.selection=1
                }
            }
        }
        dla.onChange=function(){
           if(dla.selection!=null){
               ind=ListIDget(dla.items,dla.selection.text)
               okuStr.text=dlo.items[ind]
               rubyStr.text=dlr.items[ind]
               rubyStr.active=true
           }
        }
        openBt.onClick=function(){
            new File(deepXmlPath).execute();
        };
        xmlBt.onClick=function(){
            var fileObj=new File(deepXmlPath);
            var newXmlPath=fileObj.saveDlg("新しいxmlの保存場所...","*.xml");
            if(newXmlPath){
                var selectXml=new XML("<xml><selectmode>"+ini["selectmode"]+"</selectmode></xml>");
                var oneXml=new XML("<xml><onemode>"+ini["onemode"]+"</onemode></xml>");
                var obj=new File(newXmlPath);
                if(obj.exists==false){
                    fileObj.copy(obj);
                    fileObj.remove();
                }else{
                    if(confirm("ルビ辞書を統合しますか？",ok,"ルビ辞書の統合")){
                        var send=ReadXML(newXmlPath);
                        var moto=xd;
                        delete moto["selectmode"];
                        delete moto["onemode"];
                        var childs=moto.children();
                        for(var m=0;m<childs.length();m++){
                            if(send.contains(childs[m])==false){
                                send.appendChild(childs[m]);
                            };
                        };
                        WriteText(newXmlPath,send,"UTF8");
                        xd=send;
                    }else{
                        xd=ReadXML(newXmlPath);
                    };
                };
                ini=new XML("<xd><path>"+newXmlPath+"</path></xd>");
                ini.prependChild(selectXml["selectmode"]);
                ini.prependChild(oneXml["onemode"]);
                WriteText(deepIniPath,ini,"UTF8");
                deepXmlPath=newXmlPath;
                alert("変更しました！");
                newWindow.close();
            };
        };
        selModeBt.onClick=function(){
            ini=ReadXML(deepIniPath);
            delete ini["selectmode"];
            if(selectMode=="true"){
                var selectXml=new XML("<xml><selectmode>false</selectmode></xml>");
                alert("通常ルビ振りモードに戻しました！");
            }else{
                var selectXml=new XML("<xml><selectmode>true</selectmode></xml>");
                alert("ファイル選択ルビ振りモードに変更しました！");
            };
            ini.prependChild(selectXml["selectmode"]);
            WriteText(deepIniPath,ini,"UTF8");
            newWindow.close();
        }
        oneModeBt.onClick=function(){
            ini=ReadXML(deepIniPath);
            delete ini["onemode"];
            if(oneMode=="true"){
                var oneXml=new XML("<xml><onemode>false</onemode></xml>");
                alert("全回ルビ振りモードに戻しました！");
            }else{
                var oneXml=new XML("<xml><onemode>true</onemode></xml>");
                alert("初回のみルビ振りモードに変更しました！");
            };
            ini.prependChild(oneXml["onemode"]);
            WriteText(deepIniPath,ini,"UTF8");
            newWindow.close();
        }
        delBt.onClick=function(){
            delete xd[sta3.text][ind-1]
            flag=true
            findBt.notify("onClick")
            openBt.enabled=false
            xmlBt.enabled=false
        }
        chengeBt.onClick=function(){
            if(mode=="change"){
                if(rubyStr.text==""){
                    alert("ルビ文字が入力されていません！")
                }else{
                    var tempRuby=xd[sta3.text];var TF=true//同じのがないかチェック
                    for(var se=0;se<tempRuby.length();se++){
                        if(tempRuby[se].@送り仮名.toString()==okuStr.text){
                            if(tempRuby[se].toString()==rubyStr.text){
                                TF=false
                                break;
                            }
                        }
                    }
                    if(TF){
                        if(ind==0){
                            var newRuby=new XML("<xml><"+sta3.text+">"+rubyStr.text+"</"+sta3.text+"></xml>")
                            if(okuStr.text!=""){newRuby[sta3.text].@送り仮名=okuStr.text}
                            xd.appendChild(newRuby[sta3.text])
                        }else{
                            xd[sta3.text][ind-1]=rubyStr.text
                            if(okuStr.text!=""){
                                xd[sta3.text][ind-1].@送り仮名=okuStr.text
                            }else{
                                delete xd[sta3.text][ind-1].@送り仮名
                            }
                        }
                        alert("登録しました！")
                        openBt.enabled=false
                        xmlBt.enabled=false
                        findBt.notify("onClick")
                    }else{
                        alert("既に登録されています！")
                    }
                }
            }else{
                if(rubyStr.text==""){
                    alert("ルビ文字が入力されていません！")
                }else{
                    var tempRuby=xd[sta3.text];var TF=true//同じのがないかチェック
                    for(var se=0;se<tempRuby.length();se++){
                        if(tempRuby[se].@送り仮名.toString()==okuStr.text){
                            if(tempRuby[se].toString()==rubyStr.text){
                                TF=false
                                break;
                            }
                        }
                    }
                    if(TF){
                        var newRuby=new XML("<xml><"+sta3.text+">"+rubyStr.text+"</"+sta3.text+"></xml>")
                        if(okuStr.text!=""){
                            newRuby[sta3.text].@送り仮名=okuStr.text
                        }
                        xd.appendChild(newRuby[sta3.text])
                    }
                    newWindow.close()
                }
            }
            flag=true
        }
        quitBt.onClick=function(){
            newWindow.close()
        }
        cansel.onClick=function(){
            newWindow.close()
        }
    };
    newWindow.show();
    if(flag){
        if(mode=="change"){
            WriteText(deepXmlPath,xd,"UTF8")
        }else if(mode=="new"){
            obj.rubyString=rubyStr.text
            return xd
        }else if(mode=="deep"){
            return xd
        }
    }else{
        if(mode=="new"){
            obj.rubyString=rubyStr.text
        }
        return xd
    }
}
//振られているルビから学習する
function rubiesSet(only){
    for(z=0;z<only.length;z++){
        app.findGrepPreferences.properties={findWhat:"([ぁ-ゖ]+)"};
        var tag=only[z].findGrep()[0];
        var rubies=["","",""]
        if(tag!=undefined){
            rubies[2]=tag.contents
        }else{
            rubies[2]=""
        }
        var regObj=new RegExp("(.*)"+rubies[2],"g")
        rubies[0]=(only[z].contents).replace(regObj,"$1")
        if(only[z].rubyFlag==true){
            var cc=only[z].characters
            for(var c=0;c<cc.length;c++){
                var TF=false
                rubies[1]=cc[c].rubyString
                rubies[2]=""
                if(rubies[1]!=""){
                    rubies[0]=cc[c].contents
                    for(var rr=1;rr<rubies[1].split(" ").length;rr++){
                        rubies[0]=rubies[0]+cc[c+rr].contents
                        c=c+1
                    }
                    if(c+rr<cc.length){
                        if((cc[c+rr].contents).search(new RegExp("([ぁ-ゖ]+)","g"))!=-1){
                            for(var d=c+1;d<cc.length;d++){
                                rubies[2]=rubies[2]+cc[d].contents
                            }
                        }
                    }
                    var xdr=xd[rubies[0]]
                    if(xdr.length()!=0){
                        for(var s=0;s<xdr.length();s++){
                            if(rubies[1]==xdr[s].toString()){
                                TF=true
                            }
                        }
                    }
                    if(TF==false){
                        xd=rubiDia("deep",rubies,only[z])
                    }
                }
            }
        }
    }
}   
//ワードテキストからルビを振る
function wordRubySet(wds,wdsPlus){
    for(z=0;z<wds.length;z++){
        var rubies=wdsPlus[z].contents.split("(")
        var okuriTemp=rubies[1].split(")")
        rubies[1]=okuriTemp[0]
        rubies[2]=okuriTemp[1]
        wds[z].contents=rubies[0]
        var tempRuby=xd[rubies[0]]
        switch(tempRuby.length()){
            case 0:
                xd=rubiDia("new",rubies,wds[z]);break;
            case 1:
                var setTempRuby=xd[rubies[0]].toString()
                if(setTempRuby.replace(/ /g,"")==rubies[1]){
                    wds[z].rubyString=setTempRuby
                }
                break;
            default:
                for(var s=0;s<tempRuby.length();s++){
                    var setTempRuby=tempRuby[s].toString()
                    var tempOkur=tempRuby[s].@送り仮名.toString()
                    if(setTempRuby.replace(/ /g,"")==rubies[1]){
                        if((rubies[0]+rubies[2]).search(new RegExp(tempOkur+".*?","g"))!=-1){
                            wds[z].rubyString=tempRuby[s].toString()
                        }
                    }
                }
                break;
        }
        if(wds[z].rubyString==""){
            xd=rubiDia("new",rubies,wds[z])
        }
        wds[z].rubyFlag=true
        rf=false
    }
    return xd
}
//ルビパターンxmlの場所を決定
function rubyPath(){
    if((new File(deepXmlPath)).exists==false){
        WriteText(deepXmlPath,"<xd>\n</xd>","UTF8");
        xd=new XML("<xd>\n</xd>");
    }else{
        xd=ReadXML(deepXmlPath);
    };
    if(xd.path.toString()!=""){
        deepXmlPath=xd.path.toString();
        xd=ReadXML(deepXmlPath);
    };
};
//配列と名前を渡してインデックスを取得するメソッド
function ListIDget(List,NAM){
	var ID =0;
	try{
		for(i=0;i<List.length;i++){
			if (List[i].text==NAM){
				ID = i;
			}
		}
		return ID;
	}finally{
		return ID;
	}
}
//XML読み込み
function ReadXML(path){
	var openFile=new File(path);
	openFile.open("r");
	src=openFile.read();
	return (new XML(src));
}
//テキスト書込
function WriteText(path,str,code){
	var fileObj = new File(path);
	var flag = fileObj.open("W");
	if (!flag){
		return(false); // 以後の処理はしない
	}
	fileObj.encoding = code;  // 出力する文字コードを指定
	fileObj.write(str);
	fileObj.close();
	return(true);
}
