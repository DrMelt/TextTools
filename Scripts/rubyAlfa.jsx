/*
rubyAlfa.jsx
(c)2017 Shock tm
rubyDeepLearning.jsxでxmlに蓄積したデータからルビを振る。
もしくは手で振り、それを学習する。
2016-02-04 ver1.0 公開バージョン
2019-03-14 ver1.1 グループルビに対応
2019-06-21 ver1.2 ルビ親文字のエラーに対応
2019-08-01 ver1.3 中止ボタン追加
2019-08-19 ver1.4 選択文字に対応
2019-09-07 ver1.5 ファイル選択ルビ振りモード追加・初回のみルビ振りモード追加・ルビに区切りが無く、親文字が１文字以上の場合はグループルビにする
2019-09-10 ver1.6 バグフィクス
*/
var xd;
var deepIniPath=decodeURI(Folder.temp+"/rubyDeep.xml");
var ini=ReadXML(deepIniPath);
var deepXmlPath=decodeURI(Folder.temp+"/rubyDeep.xml");
var fileMode="false";
var oneMode="false";
var firstArray=new Array();
firstArray.push("first");
rubyDeepMain();
//取り消し用
function rubyDeepMain(){
    app.doScript("doRubyDeepMain()", ScriptLanguage.JAVASCRIPT, [], UndoModes.fastEntireScript);
}
//メーンルーチン
function doRubyDeepMain(){
    try{
        rubyPath()
        if (app.documents.length != 0){
            if (app.selection.length!=0){
                fileMode=ini["selectmode"].toString();
                oneMode=ini["onemode"].toString();
                if(oneMode==""){oneMode=false};
                var moto;
                if((fileMode=="")|(fileMode=="false")){
                    fileMode="false";
                    moto=ReadXML(deepXmlPath);
                }else{
                    var fileObj=new File(deepXmlPath);
                    var newXmlPath=fileObj.openDlg("適用するルビ振りxmlの場所...","*.xml");
                    if(newXmlPath){
                        xd=ReadXML(newXmlPath);
                        moto=xd;
                    }else{
                        alert("中止しました。");
                        return;
                    };
                };
                for(i=0;i<app.selection.length;i++){
                    var selObj=app.selection[i];
                    switch(selObj.constructor.name){
                        case"TextFrame":
                            var tag=selObj.parentStory
                            app.findGrepPreferences.properties={findWhat:"~K+([ぁ-ゖ]*)"};
                            var only=tag.findGrep();
                            rubiesSet(only,fileMode)
                            break;
                        case"Text":case"TextColumn":case"TextStyleRange":case"Paragraph":case"Line":case"Word":case"Character":
                            var tag=selObj;
                            app.findGrepPreferences.properties={findWhat:"~K+([ぁ-ゖ]*)"};
                            var only=tag.findGrep();
                            rubiesSet(only,fileMode)
                            break;
                        default:
                            break;
                    };
                }
                if(xd!=moto){
                    WriteText(deepXmlPath,xd,"UTF8")
                }
            }else{
                alert("テキスト選択・テキストフレーム選択をして下さい。")
            }
        }
    }catch(e){
        alert("error:"+e,"doRubyMain")
    }
}
//ルビ登録ダイアログ
function rubiDia(rubies,obj){
    var newWindow= new Window ('dialog', "rubyAlfa1.4  writen by shock tm", [  0,  0,  0+260,  0+175])
    var flag=false;var ind=0;var end=true;
    with(newWindow){;center();
        var sta1=add('statictext', [ 10, 13, 10+240, 15+ 30],"テキスト「"+rubies[0]+"」のルビを振りますか？\n区切り部分は「欧文スペース」をどうぞ。",{multiline:true});
        var sta2=add('statictext', [ 10, 53, 10+ 60, 53+ 14],'親文字：',{multiline:false});
        var sta3=add('statictext', [ 70, 53, 70+ 55, 53+ 14],'' , {multiline:false});
        var sta4=add('statictext', [125, 48,125+ 40, 48+ 14],'候補：',{multiline:false});sta4.visible=false;
        var sta5=add('statictext', [ 10, 78, 10+ 60, 78+ 14],'ルビ文字：',{multiline:false});
        var rubyStr=add('edittext',[ 70, 75, 70+185, 75+ 20],'',{multiline:false});
        var sta6=add('statictext', [ 10,108, 10+ 60,108+ 14],'送り仮名：',{multiline:false});
        var okuStr=add('edittext', [ 70,105, 70+185,105+ 20],'',{multiline:false});
        var allCan=add('button',   [ 10,135, 10+ 55,135+ 25],'中止');
        var chengeBt=add('button', [200,135,200+ 55,135+ 25],'登録',{name: 'ok'});
        var cansel=add('button',   [140,135,140+ 55,135+ 25],'Cancel');cansel.visible=false
        cansel.visible=true
        sta3.text=rubies[0]
        rubyStr.text=rubies[1]
        okuStr.text=rubies[2]
        rubyStr.active=true
        chengeBt.onClick=function(){
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
            end=false;
            flag=true;
            newWindow.close()
        }
        cansel.onClick=function(){
            end=false;
            newWindow.close()
        }
        allCan.onClick=function(){
            newWindow.close();
        }
    };
    newWindow.show();
    if(flag){
        app.findGrepPreferences.properties={findWhat:sta3.text};
        var taget=obj.findGrep()[0];
        taget.rubyString=rubyStr.text
        taget.rubyFlag=true
        WriteText(deepXmlPath,xd,"UTF8")
    }
    return end;
}
//ルビデータベースからルビを振るもしくは手で振る
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
        if(only[z].rubyFlag){
            if(confirm("【"+rubies[0]+"】には既にルビが振られています。\n【データベース】よりルビを振り直しますか？", false, "既にルビが振られています。")){
                only[z].rubyFlag=false;
            };
        };
        if(only[z].rubyFlag==false){
            if(oneMode=="true"){
                var first=firstArray.join(",")+",end";
                if(first.indexOf(","+rubies[0]+",")!=-1){
                    continue;
                };
            };
            if(rubies[2]==""){
                var setTempRuby=xd[rubies[0]].toString()
                if(setTempRuby!=""){
                    var xdr=xd[rubies[0]]
                    for(var s=0;s<xdr.length();s++){
                        if(rubies[1]==xdr[s].@送り仮名.toString()){
                            app.findGrepPreferences.properties={findWhat:rubies[0]};
                            var target=only[z].findGrep()[0];
                            target.rubyString=xdr[s].toString()
                            target.rubyFlag=true;
                            if((xdr[s].toString().split(" ").length==1)&(target.contents.length>1)){
                                target.rubyType=RubyTypes.GROUP_RUBY;//ルビに区切りが無く、親文字が１文字以上の場合はグループルビにする
                            };
                        };
                    };
                };
            }else{
                var xdr=xd[rubies[0]]
                for(var s=0;s<xdr.length();s++){
                    if(only[z].rubyFlag==false){
                        switch(xdr.length()){
                            case 0:break;
                            default:
                                var xdTempRuby=xdr[s].toString()
                                var xdOkuri=xdr[s].@送り仮名.toString()
                                if((rubies[0]+rubies[2]).search(new RegExp(xdOkuri+".*?","g"))!=-1){
                                    app.findGrepPreferences.properties={findWhat:rubies[0]};
                                    var target=only[z].findGrep()[0];
                                    target.rubyString=xdTempRuby
                                    target.rubyFlag=true
                                    if((xdr[s].toString().split(" ").length==1)&(target.contents.length>1)){
                                        target.rubyType=RubyTypes.GROUP_RUBY;
                                    };
                                }
                            break;
                        }
                    }else{
                        break;
                    }
                }
            }
        }
        if((only[z].rubyFlag==false)&(fileMode=="false")){
            var bool=rubiDia(rubies,only[z]);
            if(bool){return}
        };
        if(only[z].rubyFlag){
            firstArray.push(rubies[0]);
        };
    };
};
//ルビパターンxmlの場所を決定
function rubyPath(){
    if((new File(deepXmlPath)).exists==false){
        WriteText(deepXmlPath,"<xd>\n</xd>","UTF8")
        xd=new XML("<xd>\n</xd>")
    }else{
        xd=ReadXML(deepXmlPath)
    }
    if(xd.path.toString()!=""){
        deepXmlPath=xd.path.toString()
        xd=ReadXML(deepXmlPath)
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
