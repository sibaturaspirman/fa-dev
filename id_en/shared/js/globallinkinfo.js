// グローバルドキュメントリンク
var GET_DOCUMENT_LINK_INFO = "https://www.mitsubishielectric.com/app/fa/DocumentSearchService/GetDocumentLinkInfo.do?";
var NEW_GET_DOCUMENT_LINK_INFO = "https://www.mitsubishielectric.com/fa/linkinfo/DocumentSearchService/GetDocumentLinkInfoJson?";
// ドキュメントのサーバ
var DOC_DOMAIN_NAME="";
// イメージのサーバー
var IMG_DOMAIN_NAME="https://www.mitsubishielectric.co.jp";

var DOC_NO = "\"documentNO1\"";
var DOC_NO2 = "\"documentNO\"";
var KISYU_NO = "\"kisyuNO\"";
var DOCUMENT_TYPE = "\"documentType\"";
var SHIRYO_ID = "\"shiryoID\"";
var SEQ_NO = "\"seqNo\"";

var PDF_ICON_IMG="/fa/cmn11/grparts/pdficon.gif";
var EMA_ICON_IMG="";
var EPUB_ICON_IMG="";
var DNP_ICON_IMG="";


var DIRECT_LINK_CATALOG  ="/app/fa/download/search.do?mode=catalog&kisyu=";
//var DIRECT_LINK_TECHNEWS ="/app/fa/products/documentsearch/technews/technicalNews.do?kisyu=";
var DIRECT_LINK_TECHNEWS ="/app/fa/download/techinfo/search.do?mode=technews&kisyu=";
//var DIRECT_LINK_SALESE   ="/app/fa/products/documentsearch/sales/salesService.do?kisyu=";
var DIRECT_LINK_SALESE   ="/app/fa/download/techinfo/search.do?mode=sales&kisyu=";
//var DIRECT_LINK_NPRODUCT ="/app/fa/products/documentsearch/nproduct/newProductNews.do?kisyu=";
var DIRECT_LINK_NPRODUCT ="/app/fa/download/search.do?mode=nproduct&kisyu=";
var DIRECT_LINK_MANUAL   ="/app/fa/download/search.do?mode=manual&kisyu=";
var DIRECT_LINK_CURVE     ="/app/fa/download/techinfo/search.do?mode=curve&kisyu=";
var DIRECT_LINK_TECHSHEET ="/app/fa/download/techinfo/search.do?mode=techsheet&kisyu=";
var DIRECT_LINK_REPORT    ="/app/fa/download/techinfo/search.do?mode=report&kisyu=";
var DIRECT_LINK_TECHLIB    ="/app/fa/download/techinfo/search.do?mode=techlib&kisyu=";

var DIRECT_LINK_CATALOG_NEW  ="/fa/download/search.do?mode=catalog&kisyu=";
var DIRECT_LINK_TECHNEWS_NEW ="/fa/download/techinfo/search.do?mode=technews&kisyu=";
var DIRECT_LINK_SALESE_NEW   ="/fa/download/techinfo/search.do?mode=sales&kisyu=";
var DIRECT_LINK_NPRODUCT_NEW ="/fa/download/search.do?mode=nproduct&kisyu=";
var DIRECT_LINK_MANUAL_NEW   ="/fa/download/search.do?mode=manual&kisyu=";
var DIRECT_LINK_CURVE_NEW     ="/fa/download/techinfo/search.do?mode=curve&kisyu=";
var DIRECT_LINK_TECHSHEET_NEW ="/fa/download/techinfo/search.do?mode=techsheet&kisyu=";
var DIRECT_LINK_REPORT_NEW    ="/fa/download/techinfo/search.do?mode=report&kisyu=";
var DIRECT_LINK_TECHLIB_NEW    ="/fa/download/techinfo/search.do?mode=techlib&kisyu=";

// 要素番号
var xmlno = 0;
var xmldispno = 1;
var xmlurl = 2;
var xmlver = 3;
var xmltitle = 4;
var xmlimgKisyu = 5;
var xmlimgCatalog = 6;
var xmlimgOption = 7;
var xmldate = 8;
var xmlshiryoStatus= 9;
var xmlshiryoDel=10;
var xmlanchor=11;
var xmlkisyupath =12;
var xmlbunruil=13;
var xmlbunruis=14;
var xmlpdfsize=15;
var xmlemaurl = 16;
var xmlemaver = 17;
var xmlemasize=18;
var xmlepuburl = 19;
var xmlepubver = 20;
var xmlepubsize=21;
var xmldnpurl = 22;
var xmldnpver = 23;
var xmldnpsize=24;

var count=0;
function createXmlHttp(){
 if(window.XMLHttpRequest){
  // IE7以上、SAFARI、FireFox
  return new XMLHttpRequest();
 } else if(window.ActiveXObject){
  // IE5、IE6
  try {
   // MSXML3
   return new ActiveXObject("Msxml2.XMLHTTP");
  } catch(e) {
   // MSXML2
   return new ActiveXObject("Microsoft.XMLHTTP");
  }
 } else{
  // 非対応のブラウザ
  return null;
 }
}

function loadDataLinkinfo() {
	$("<script type='text/javascript' charset='utf-8' src='/fa/shared/js/ikouKisyu.js'></script>").appendTo("body");
}


// カタログリンク表示
function toCatalogInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid,docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'catalog', DIRECT_LINK_CATALOG, DIRECT_LINK_CATALOG_NEW);
         // グローバルの場合は"9","11"は対象外 グローバルドキュメントリンク対応で対象にする
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
         //if (anchorType == "10") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}
// カタログリンク表示
function toCatalogInfoA(elemid, kisyuNo, anchorType, titleType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if( kisyuNo != null && kisyuNo != "" && anchorType != null && anchorType != "" && titleType != null && titleType != ""){
         var ret = makeLinkPage2(kisyuNo, anchorType, titleType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'cataloga', DIRECT_LINK_CATALOG, DIRECT_LINK_CATALOG_NEW);
         // グローバルの場合は"9","11"は対象外
         //if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
         if (anchorType == "10") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}

//テクニカルニュースリンク表示
function toTechnewsInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'technews', DIRECT_LINK_TECHNEWS, DIRECT_LINK_TECHNEWS_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}
// セールスとサービスリンク表示
function toSalesInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'sales', DIRECT_LINK_SALESE, DIRECT_LINK_SALESE_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
    }
}
// 新製品ニュースリンク表示
function toNproductInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'nproduct', DIRECT_LINK_NPRODUCT, DIRECT_LINK_NPRODUCT_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
    }
}
// マニュアルリンク表示
function toManualInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
        var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'manual', DIRECT_LINK_MANUAL, DIRECT_LINK_MANUAL_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
         //if (anchorType == "10") {
          if (ret != null && ret != "") {
             return ret;
          } else {
              return "";
          }
        }
    } else if (docNo != null && docNo != "") {
        var ret = makeLinkPageOrgManual(docNo,kisyuNo);
         if (anchorType == "10" || anchorType == "11") {
         //if (anchorType == "10") {
          if (ret != null && ret != "") {
             return ret;
          } else {
              return "";
          }
        }
    }
}
// マニュアルリンク表示 マイページ用
function toManualInfoA(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
        var obj = document.getElementById(kisyuNo);
        if (obj != null) {
            makeLinkPageOrgManualA(docNo,kisyuNo);
            return;
        }
        var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'manual', DIRECT_LINK_MANUAL, DIRECT_LINK_MANUAL_NEW);

         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
         //if (anchorType == "10") {
          if (ret != null && ret != "") {
             return ret;
          } else {
             return "";
          }
        }
    }
}
// その他のPDF情報リンク表示
function toOthersInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
        var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'other' ,'' ,'');
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
          if (ret != null && ret != "") {
             return ret;
          } else {
              return "";
          }
        }
    }
}
// テクニカルライブラリリンク表示
//function toTecLibInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
//    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
//         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'techlib');
//         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
//           if (ret != null && ret != "") {
//              return ret;
//           } else {
//              return "";
//           }
//         }
//    }
//}
//特性曲線データ表示
function toCurveInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'curve', DIRECT_LINK_CURVE, DIRECT_LINK_CURVE_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}
//テクニカルシート表示
function toTechsheetInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'techsheet', DIRECT_LINK_TECHSHEET, DIRECT_LINK_TECHSHEET_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}
//テストレポートリンク表示
function toReportInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'report', DIRECT_LINK_REPORT, DIRECT_LINK_REPORT_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}
//技術資料表示
function toTechLibInfo(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol){
    if(docNo != null && docNo != "" && kisyuNo != null && kisyuNo != ""){
         var ret = makeLinkPage(elemid, docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, 'techlib', DIRECT_LINK_TECHLIB, DIRECT_LINK_TECHLIB_NEW);
         if (anchorType == "9" || anchorType == "10" || anchorType == "11") {
           if (ret != null && ret != "") {
              return ret;
           } else {
              return "";
           }
         }
     }
}
//******************************************************************************************************************************************************
// ドキュメントリンク HTMLイメージ作成
//******************************************************************************************************************************************************
function makeLinkPage(elemid,docNo, kisyuNo, anchorType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, docType, directlink, directlinkNew) {
	loadDataLinkinfo();
	if (anchorType < 0 || anchorType > 17) {
		document.write(errOut(anchorType, anchorTitle,docNo));
		return;
	}
    // ドキュメントリンクへの通信
    // パラメータの設定
    if (docNo == "" && pdfVol != null && pdfVol != "") {
		var param = { "kisyuNO":kisyuNo, "documentType":docType, "seqNo":pdfVol};
    } else if (docNo == "" && (pdfVol == null || pdfVol == "")) {
		var param = { "kisyuNO":kisyuNo, "documentType":docType};
    } else if (docNo != "" && pdfVol != null && pdfVol != "") {
		var param = { "documentNO1": docNo, "kisyuNO":kisyuNo, "documentType":docType, "seqNo":pdfVol};
    } else {
		var param = { "documentNO1": docNo, "kisyuNO":kisyuNo, "documentType":docType};
    }
//    var localCount = count;
//	document.write("<div class=\"test"+localCount+"\"></div>");
//	for (i=0; i< 2; i++) {
    
	// 移行対象機種であるか判定
	var count = 0;
	var flg = false;
	for ( count=0; count<=KISYU_ARRAY.length; count++) {
		if (KISYU_ARRAY[count] == kisyuNo) {
			flg = true;
			break;
		}
	}
	// 機種名での判定(機種IDでフラグがtrueになった場合はスキップ）
	if (!flg) {
		for ( count=0; count<=KISYU_NAME_ARRAY.length; count++) {
			if (KISYU_NAME_ARRAY[count] == kisyuNo) {
				flg = true;
				break;
			}
		}
	}
	
	var postUrl = "";
    // 通信をOPEN
    if (flg) {
        postUrl = NEW_GET_DOCUMENT_LINK_INFO;
        directlink = directlinkNew
    } else {
        postUrl = GET_DOCUMENT_LINK_INFO;
    }

		// 通信をOPEN
		$.ajax({
			type: "GET",
			cache: true,
			url:postUrl,
			async:false,
			data: param,
			dataType: 'jsonp'}).done( function (json){
				// 返り値ＸＭＬの取得
				var chtm = new Array();
				var errorcode = json.ErrorCD;
				if(errorcode == 0){
					// DocumentLinkInfoが複数存在した場合は1つ目のデータを使う
					var localDocumentLinkInfo = json.DocumentLinkInfoList.DocumentLinkInfo;
					if (Object.prototype.toString.call(localDocumentLinkInfo) === '[object Array]') {
						localDocumentLinkInfo = json.DocumentLinkInfoList.DocumentLinkInfo[0];
					}
					datas = ansXmlDataToArray(localDocumentLinkInfo);
					var dispStatus = dispStatusSet(datas);
					var directurl = directlink;
					var pdfurl = pdfurlSet(datas,directFlg);
					var documentNo = documentNoSet(datas);
					var pdfver = versionSet(datas,directFlg);

					
					// ダイレクトURL設定
					if(directurl != null && directurl != ""){
						directurl = directurl + datas[xmlkisyupath];
					}
					// タイトルタイプ
					if ( (Number(anchorType) >= 0 && Number(anchorType) < 5) || anchorType == "11" ) {
					//if ( (Number(anchorType) >= 0 && Number(anchorType) < 5) ) {
						chtm.push("<table>");
						chtm.push(titleSet(anchorType, documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
						if(anchorIssueFlag == "1"){
							setDate ="";
							if (datas[xmldate]) { setDate = datas[xmldate]; }
							chtm.push("<tr><td class='docLinkDate'>" + setDate +"</td></tr>");
						}
						chtm.push("</table>");
					}else if(anchorType == "5"){
						chtm.push("<table>");
						chtm.push(imageSet(documentNo, datas[xmlimgKisyu], directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
						chtm.push("</table>");
					}else if(anchorType == "6"){
						// 画像（カタログ一覧用）
						chtm.push("<table>");
						chtm.push(imageSet(documentNo, datas[xmlimgCatalog], directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
						chtm.push("</table>");
					// グローバルの場合は"7"～"9"は対象外 →グローバルドキュメントリンク対応で対象に修正
					}else if(anchorType == "7") {
						// 画像（任意）
						chtm.push(imageSet(documentNo, datas[xmlimgOption], directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
					}else if(anchorType == "8") {
						// 画像（PDFアイコン）
						var imgUrl = ""
						if (directFlg == 3){
							imgUrl=EMA_ICON_IMG;
						} else if (directFlg == 4) {
							imgUrl=EPUB_ICON_IMG;
						} else if (directFlg == 5) {
							imgUrl=DNP_ICON_IMG;
						} else {
							imgUrl=PDF_ICON_IMG;
						}
						chtm.push(imageSet(documentNo, imgUrl, directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
					}else if(anchorType == "9") {
						// ページジャンプ
						return pageLocation(directFlg, pdfurl, directurl, datas[xmlno], pdfPage);
					}else if(anchorType == "10") {
						// ページジャンプ
						chtm.push(getUrl(directFlg, pdfurl, directurl, pdfPage));
					// グローバルの場合は"12"は対象外 →グローバルドキュメントリンク対応で対象に修正
					}else if(anchorType == "12") {
						uptoolTopics(anchorType, documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage);
						return;
					}else if(anchorType == "13") {
						// 発行年月のみ
						if (datas[xmldate] != null && datas[xmldate] != "") {
							titleOut(anchorType, documentNo, pdfver, datas[xmldate], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage);
							return;
							//chtm.push(datas[xmldate]);
						} else {
							chtm.push(errOut(anchorType, anchorTitle,docNo));
						}
					}else if(anchorType == "14") {
						// サイズのみ
						var filesize = sizeSet(datas,directFlg);
						if (filesize != null && filesize != "") {
							var orgsize = Number(filesize);
							orgsize = orgsize / 1024 / 1024;
							var setsize = "" +orgsize;
							if ( setsize.indexOf(".") != -1) {
							   var splitsize = setsize.split(".");
							   setsize = splitsize[0] + "." + splitsize[1].substring(0,1);
							} else {
							   setsize = setsize + ".0";
							}
							if (setsize == "0.0") {
							   setsize = "0.1";
							}
							var ssetsize = setsize + "MB";
							titleOut(anchorType, documentNo, pdfver, ssetsize, anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage);
							return;
							//chtm.push(setsize + "MB");
						} else {
							chtm.push(errOut(anchorType, anchorTitle,docNo));
						}
					}else if(anchorType == "15") {
						// 分類のみ
						if (datas[xmlbunruil] != null && datas[xmlbunruil] != "") {
							titleOut(anchorType, documentNo, pdfver, datas[xmlbunruil], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage);
							return;
							//chtm.push(datas[xmlbunruil]);
						} else {
							chtm.push(errOut(anchorType, anchorTitle,docNo));
						}
					}else if(anchorType == "16") {
						docNoOut(anchorType, documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage);
						return;
					}else if(anchorType == "17") {
						titleOut(anchorType, documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage);
						return;
					}else {
						// 分割
						var outhtm = new Array();
						var split = anchorType.split(" ");
						if (split.length) {
							for (i=0; i< split.length; i++) {
								if (isFinite(split[i])) {
									switch(Number(split[i])) {
									case 0:
									case 1:
									case 2:
									case 3:
										outhtm.push(titleSetNoTable(split[i], documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, directFlg, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
										break;
									case 14:
										var filesize = sizeSet(datas,directFlg);
										if (filesize != null && filesize != "") {
											var orgsize = Number(filesize);
											orgsize = orgsize / 1024 / 1024;
											var setsize = "" +orgsize;
											if ( setsize.indexOf(".") != -1) {
												var splitsize = setsize.split(".");
												setsize = splitsize[0] + "." + splitsize[1].substring(0,1);
											} else {
												setsize = setsize + ".0";
											}
											if (setsize == "0.0") {
												setsize = "0.1";
											}
											outhtm.push(setsize + "MB");
										} else {
											outhtm.push(errOut(anchorType, anchorTitle,docNo));
										}
										break;
									case 15:
										// 分類のみ
										if (datas[xmlbunruil] != null && datas[xmlbunruil] != "") {
											outhtm.push(datas[xmlbunruil]);
										} else {
											outhtm.push(errOut(anchorType, anchorTitle,docNo));
										}
										break;
									case 18:
										// 資料名称のみ
										outhtm.push(titleSetNoTable("0", documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, 0, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
										break;
									case 19:
										// 資料名称のみ
										outhtm.push(titleSetNoTable("1", documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, 0, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
										break;
									case 20:
										// 資料名称のみ
										outhtm.push(titleSetNoTable("2", documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, 0, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
										break;
									case 21:
										// 資料名称のみ
										outhtm.push(titleSetNoTable("3", documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, 0, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
										break;
									case 22:
										// 資料名称のみ
										outhtm.push(titleSetNoTable("4", documentNo, pdfver, datas[xmltitle], anchorTitle, anchorNewLine, 0, pdfurl, directurl,datas[xmlanchor], datas[xmlno], pdfPage));
										break;
									}
								} else {
									outhtm.push(split[i]);
								}
							}
							while(!$('#'+elemid)) {
							}
							$('#'+elemid).html(outhtm.join(""));
							//document.write(outhtm.join(""));
							return;
						}
					}
					// ステータス
					chtm.push(dispStatus);
					if (anchorType == "11") {
					   return chtm.join("");
					} else {
						while(!$('#'+elemid)) {
						}
						$('#'+elemid).html(chtm.join(""));
					//   document.write(chtm.join(""));
					}
					
					return;
				} else {
					if (anchorType == "10" || anchorType == "11") {
					 return;
					}
					chtm.push(errOut(anchorType, anchorTitle,docNo));
					while(!$('#'+elemid)) {
					}
					$('#'+elemid).html(chtm.join(""));
				}
//				break;
		});
//	}
}

function titleSetNoTable(titleType, docNo, docVer, title, anchorTitle, anchorNewLine, directFlg, pdfUrl, directUrl, anchorUrl, no, pdfPage) {
	if (!titleType) { titleType = ""; }
	if (!docNo) { docNo = ""; }
	if (!docVer) { docVer = ""; }
	if (!title) { title = ""; }
	if (!anchorTitle) { anchorTitle = ""; }
	if (!anchorNewLine) { anchorNewLine = ""; }
	if (!directFlg) { directFlg = ""; }
	if (!pdfUrl) { pdfUrl = ""; }
	if (!directUrl) { directUrl = ""; }
	if (!anchorUrl) { anchorUrl = ""; }
	if (!no) { no = ""; }
	if (!pdfPage) { pdfPage = ""; }
	var titlehtm = new Array();
	// タイトルタイプ
	// ダイレクト表示フラグ
	if(directFlg == 0){
	}else if(directFlg == 1){
		var setPdfUrl = pdfUrl;
		if (pdfPage != null && pdfPage != "") {
			setPdfUrl = pdfUrl + "#page="+pdfPage;
		}
		titlehtm.push("<a href='" + setPdfUrl + "' title='" + docNo + "' target='_blank'>");
	} else if(directFlg == 2 && directUrl != "") {
		titlehtm.push("<a href='" + directUrl + "#" + no + "' title='" + docNo + "' target='_top'>");
	} else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
		titlehtm.push("<a href='" + pdfUrl +  "' title='" + docNo + "' target='_blank'>");
	} else {
	}
	if(titleType == "0" || titleType == "11"){
	 //資料番号 + "-" + バージョン
		if (docVer != null && docVer != "") {
			titlehtm.push(docNo + "-" + docVer);
		} else {
			titlehtm.push(docNo);
		}
	}else if(titleType == "1"){
		//資料番号 + バージョン
		if (docVer != null && docVer != "") {
			titlehtm.push(docNo + docVer);
		} else {
			titlehtm.push(docNo);
		}
	}else if(titleType == "2"){
		//資料番号
		titlehtm.push(docNo);
	}else if(titleType == "3"){
		//資料名称
		var setTitle;
		if (anchorNewLine != null && anchorNewLine != "") {
		 var buf1 = title.substring(0, anchorNewLine);
		 var buf2 = title.substring(anchorNewLine);
		 setTitle = buf1+"<br>"+buf2;
		} else {
		 setTitle =title;
		}
		titlehtm.push(setTitle);
	}else if(titleType == "4"){
		//任意文字を設定
		if (anchorTitle == '') {
		  titlehtm.push(docNo);
		} else {
		  titlehtm.push(anchorTitle);
		}
	} else {
		return '';
	}
	// ダイレクト表示フラグ
	if(directFlg == 0){
	}else if(directFlg == 1){
		titlehtm.push("</a>");
	}else if(directFlg == 2 && directUrl != ""){
		titlehtm.push("</a>");
	} else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
		titlehtm.push("</a>");
	}else{
	}
	return titlehtm.join("");
}


//******************************************************************************************************************************************************
// ドキュメントリンク おすすめカタログ
//******************************************************************************************************************************************************
function makeLinkPage2(kisyuNo, anchorType, titleType, directFlg, anchorIssueFlag, anchorNewLine, anchorTitle, imgWidth, imgHeight, imgBorder, pdfPage, pdfVol, docType, directlink) {
	loadDataLinkinfo();
	if (anchorType < 0 || anchorType > 15) {
		document.write(errOut(anchorType, anchorTitle,docNo));
		return;
	}
    // ドキュメントリンクへの通信
    // パラメータの設定
    var parameter = KISYU_NO + "=" + kisyuNo + "&" +DOCUMENT_TYPE + "=" + docType;
    var httpObj = createXmlHttp();
    if(httpObj == null){
        return;
    }
    for (i=0; i< 2; i++) {
	// 移行対象機種であるか判定
	var count = 0;
	var flg = false;
	for ( count=0; count<=KISYU_ARRAY.length; count++) {
		if (KISYU_ARRAY[count] == kisyuNo) {
			flg = true;
			break;
		}
	}
	// 機種名での判定(機種IDでフラグがtrueになった場合はスキップ）
	if (!flg) {
		for ( count=0; count<=KISYU_NAME_ARRAY.length; count++) {
			if (KISYU_NAME_ARRAY[count] == kisyuNo) {
				flg = true;
				break;
			}
		}
	}

    // 通信をOPEN
    if (flg) {
        httpObj.open("POST", NEW_GET_DOCUMENT_LINK_INFO, false);
    } else {
        httpObj.open("POST", GET_DOCUMENT_LINK_INFO, false);
    }
    httpObj.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    httpObj.send(parameter);
    if(httpObj.readyState == 4){
        if(httpObj.status == 200){
             // 返り値ＸＭＬの取得
             xmlData  = httpObj.responseXML;
             var errorcode = xmlData.getElementsByTagName("ErrorCD")[0].childNodes[0].nodeValue;
             if(errorcode == 0){
               var datas = new Array();
               datas = ansXmlDataToArray(xmlData);

               var dispStatus = dispStatusSet(datas);
               var directurl = directlink;
               var pdfurl = pdfurlSet(datas,directFlg);
               var documentNo = documentNoSet(datas);

               var chtm = new Array();
               chtm.push("<table width='100%'>");
               // ダイレクトURL設定
               if(directurl != null && directurl != ""){
                 directurl = directurl + datas[xmlkisyupath];
               }
               // 画像タイプ               
               if(anchorType == "5"){
                   // 画像（機種トップ用）
                   chtm.push(imageSet(documentNo, datas[xmlimgKisyu], directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
               }else if(anchorType == "6"){
                   // 画像（カタログ一覧用）
                   chtm.push(imageSet(documentNo, datas[xmlimgCatalog], directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
               // グローバルの場合は"7"～"9"は対象外
               //}else if(anchorType == "7") {
               //    // 画像（任意）
               //    chtm.push(imageSet(documentNo, datas[xmlimgOption], directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
               //}else if(anchorType == "8") {
               //    // 画像（PDFアイコン）
               //    var imgUrl = ""
               //    if (directFlg == 3){
               //        imgUrl=EMA_ICON_IMG;
               //    } else if (directFlg == 4) {
               //        imgUrl=EPUB_ICON_IMG;
               //    } else if (directFlg == 5) {
               //        imgUrl=DNP_ICON_IMG;
               //    } else {
               //        imgUrl=PDF_ICON_IMG;
               //   }
               //    chtm.push(imageSet(documentNo, imgUrl, directFlg, pdfurl, directurl, datas[xmlanchor], imgWidth, imgHeight, datas[xmlno], imgBorder, pdfPage));
               //}else if(anchorType == "9") {
               //    // ページジャンプ
               //    return pageLocation(directFlg, pdfurl, directurl, datas[xmlno], pdfPage);
               }

               // タイトルタイプ
               chtm.push(titleSet(titleType, documentNo, datas[xmlver], datas[xmltitle], anchorTitle, anchorNewLine,0,'','','', datas[xmlno], pdfPage));

               // 発行(改定) 日付フラグ
               if (anchorIssueFlag == "1") {
                   chtm.push("<tr><td class='docLinkDate'>");
                   setDate ="";
                   if (datas[xmldate]) { setDate = datas[xmldate]; }
                   chtm.push("[" +setDate+ "]");
                   chtm.push("</td></tr>");
               }
               // ステータス
               chtm.push(dispStatus);
               chtm.push("</table>");
               // 表示処理
               document.write(chtm.join(""));
               return;
             } else {
               document.write(errOut(titleType, anchorTitle));
             }
             break;
         }else{
         // 通信エラー
           if (i > 0) {
             document.write(errOut(titleType, anchorTitle));
           }
         }
     }
     }
}

//******************************************************************************************************************************************************
// マニュアルリンク（引数１）
//******************************************************************************************************************************************************
function makeLinkPageOrgManual(docNo,kisyuNo) {
	loadDataLinkinfo();

    var NO_DATA = "マニュアル番号";

    // ドキュメントリンクへの通信
    // パラメータの設定
    var parameter = DOC_NO + "=" + docNo + "&" +DOCUMENT_TYPE + "=manual";
    var httpObj = createXmlHttp();
    if(httpObj == null){
        return;
    }
    for (i=0; i< 2; i++) {
	// 移行対象機種であるか判定
	var count = 0;
	var flg = false;
	for ( count=0; count<=KISYU_ARRAY.length; count++) {
		if (KISYU_ARRAY[count] == kisyuNo) {
			flg = true;
			break;
		}
	}
	// 機種名での判定(機種IDでフラグがtrueになった場合はスキップ）
	if (!flg) {
		for ( count=0; count<=KISYU_NAME_ARRAY.length; count++) {
			if (KISYU_NAME_ARRAY[count] == kisyuNo) {
				flg = true;
				break;
			}
		}
	}

    // 通信をOPEN
    if (flg) {
        httpObj.open("POST", NEW_GET_DOCUMENT_LINK_INFO, false);
    } else {
        httpObj.open("POST", GET_DOCUMENT_LINK_INFO, false);
    }
    httpObj.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    httpObj.send(parameter);
    if(httpObj.readyState == 4){
        if(httpObj.status == 200){
             // 返り値ＸＭＬの取得
             xmlData  = httpObj.responseXML;
             var errorcode = xmlData.getElementsByTagName("ErrorCD")[0].childNodes[0].nodeValue;
             if(errorcode == 0){
               var chtm = new Array();
               datas = ansXmlDataToArray(xmlData);
               var dispStatus = dispStatusSet(datas);
               var pdfurl = pdfurlSet(datas,directFlg);
               var version = versionSet(datas,directFlg);
               var documentNo = documentNoSet(datas);
               if(pdfurl != null && pdfurl != "" && documentNo != null &&documentNo != "") {
                 chtm.push("<a href=\"");
                 chtm.push(pdfurl);
                 chtm.push("\">");
                 chtm.push(documentNo);
                 chtm.push("-");
                 chtm.push(version);
                 chtm.push("</a>");
                 document.write(chtm.join(""));
                 return;
               }
             }
         }else{
         // 通信エラー
         }
     }
     }
     chtm.push(docNo);
     document.write(chtm.join(""));
}

//******************************************************************************************************************************************************
// マニュアルリンク（引数２）
//******************************************************************************************************************************************************
function makeLinkPageOrgManualA(docNo, id) {

	loadDataLinkinfo();
    var NO_DATA = "マニュアル番号";
    var obj = document.getElementById(id);

    // ドキュメントリンクへの通信
    // パラメータの設定
    var parameter = DOC_NO + "=" + docNo + "&" +DOCUMENT_TYPE + "=manual";
    var httpObj = createXmlHttp();
    if(httpObj == null){
        return;
    }
    for (i=0; i< 2; i++) {
	// 移行対象機種であるか判定
	var count = 0;
	var flg = false;
	for ( count=0; count<=KISYU_ARRAY.length; count++) {
		if (KISYU_ARRAY[count] == kisyuNo) {
			flg = true;
			break;
		}
	}
	// 機種名での判定(機種IDでフラグがtrueになった場合はスキップ）
	if (!flg) {
		for ( count=0; count<=KISYU_NAME_ARRAY.length; count++) {
			if (KISYU_NAME_ARRAY[count] == kisyuNo) {
				flg = true;
				break;
			}
		}
	}

    // 通信をOPEN
    if (flg) {
        httpObj.open("POST", NEW_GET_DOCUMENT_LINK_INFO, false);
    } else {
        httpObj.open("POST", GET_DOCUMENT_LINK_INFO, false);
    }
    httpObj.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    httpObj.send(parameter);
    if(httpObj.readyState == 4){
        if(httpObj.status == 200){
             // 返り値ＸＭＬの取得
             xmlData  = httpObj.responseXML;
             var errorcode = xmlData.getElementsByTagName("ErrorCD")[0].childNodes[0].nodeValue;
             if(errorcode == 0){
               var chtm = new Array();
               datas = ansXmlDataToArray(xmlData);
               var dispStatus = dispStatusSet(datas);
               var pdfurl = pdfurlSet(datas,directFlg);
               var version = versionSet(datas,directFlg);
               var documentNo = documentNoSet(datas);
               if(pdfurl != null && pdfurl != "" && documentNo != null &&documentNo != "") {
                 obj.innerHTML = "<a href=\"" + pdfurl + "\">" + documentNo + "-" + version + "</a>";
                 return;
               }
             }
         }else{
         // 通信エラー
         }
     }
     }
    obj.innerHTML = NO_DATA;
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// 応答XMLデータを変数に格納します
// 引数 xmlData：応答XMLデータ
//---------------------------------------------
function ansXmlDataToArray(xmlData) {
   var ansXmlData = new Array();

   if(xmlData.DocAnchor != null){
     ansXmlData[xmlanchor] = xmlData.DocAnchor;
   }
   
   if(xmlData.DocNumber != null ){
     ansXmlData[xmlno] = xmlData.DocNumber;
   }
   
   if(xmlData.DocDispNumber != null ){
     ansXmlData[xmldispno] = xmlData.DocDispNumber;
   }
   
   if(xmlData.DocKisyuPath != null){
     ansXmlData[xmlkisyupath] = xmlData.DocKisyuPath;
   }
   
   if(xmlData.DocPdfURL != null){
     var membersurl = xmlData.DocPdfURL;
     var nonmenbersurl = membersurl.replace("/members","");
     ansXmlData[xmlurl] = nonmenbersurl;
   }
   
   if (xmlData.DocPdfVer != null) {
       ansXmlData[xmlver] = xmlData.DocPdfVer;
   }
   
   if(xmlData.DocTitle != null){
     ansXmlData[xmltitle] = xmlData.DocTitle;
   }
   
   if(xmlData.DocImgKisyuURL != null){
     ansXmlData[xmlimgKisyu] = xmlData.DocImgKisyuURL;
   }
   if(xmlData.DocImgCatalogURL != null){
     ansXmlData[xmlimgCatalog] = xmlData.DocImgCatalogURL;
   }
   if(xmlData.DocImgOptionURL != null){
     ansXmlData[xmlimgOption] = xmlData.DocImgOptionURL;
   }
   if(xmlData.DocDate != null){
     var date = xmlData.DocDate;
     var buff = date.split("/");
     //ansXmlData[xmldate] = buff[0] + "年" + buff[1] + "月";
     //YYYY/MM・ﾖ・ﾏ・X
     ansXmlData[xmldate] = buff[0] + "-" + buff[1];
   }
   if(xmlData.DocStatus != null){
     ansXmlData[xmlshiryoStatus] = xmlData.DocStatus;
   }
   if(xmlData.DocRegistType != null){
     ansXmlData[xmlshiryoDel] = xmlData.DocRegistType;
   }
   if(xmlData.DocBunruiL != null){
     ansXmlData[xmlbunruil] = xmlData.DocBunruiL;
   }
   if(xmlData.DocBunruiS != null){
     ansXmlData[xmlbunruis] = xmlData.DocBunruiS;
   }
   if(xmlData.DocPdfSize != null){
     ansXmlData[xmlpdfsize] = xmlData.DocPdfSize;
   }
   // ドキュメントDB改善　フェーズ2対応
   // E-manual
   if(xmlData.DocEmaURL != null){
       ansXmlData[xmlemaurl] = xmlData.DocEmaURL;
   }
   
   if (xmlData.DocEmaVer != null) {
       ansXmlData[xmlemaver] = xmlData.DocEmaVer;
   }

   if(xmlData.DocEmaSize != null){
     ansXmlData[xmlemasize] = xmlData.DocEmaSize;
   }

	// epub
   if(xmlData.DocEpubURL != null){
       ansXmlData[xmlepuburl] = xmlData.DocEpubURL;
   }
   
   if (xmlData.DocEpubVer != null) {
       ansXmlData[xmlepubver] = xmlData.DocEpubVer;
   }

   if(xmlData.DocEpubSize != null){
     ansXmlData[xmlepubsize] = xmlData.DocEpubSize;
   }

	//dnp
   if(xmlData.DocDnpURL != null){
       ansXmlData[xmldnpurl] = xmlData.DocDnpURL;
   }
   
   if (xmlData.DocDnpVer != null) {
       ansXmlData[xmldnpver] = xmlData.DocDnpVer;
   }
   
   if(xmlData.DocDnpSize != null){
     ansXmlData[xmldnpsize] = xmlData.DocDnpSize;
   }
   

   return ansXmlData;
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// 登録状態をHTMLイメージを作成します。
// 引数 datas：応答XMLデータ配列
// 戻り値　　；登録状態HTMLイメージ
//---------------------------------------------
function dispStatusSet(datas) {
   var dispStatus = new Array();
   if (datas[xmlshiryoStatus] != null && datas[xmlshiryoStatus] != "") {
       dispStatus.push("<tr>");
       dispStatus.push("<td class='docstatus'><span class='docLinkStatus'>");
       dispStatus.push(datas[xmlshiryoStatus]);
       if (datas[xmlshiryoDel] != null && datas[xmlshiryoDel] != "") {
           dispStatus.push("&nbsp;<span class='docLinkDel'>(");
           dispStatus.push(datas[xmlshiryoDel]);
           dispStatus.push(")</span>");
       }
       dispStatus.push("</span></td></tr>");
   }
   return dispStatus.join("");
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// 資料番号の作成を行います。
//---------------------------------------------
function documentNoSet(datas) {
   var documentNo = '';
   if (datas[xmldispno] != null) {
       documentNo = datas[xmldispno];
   } else {
       documentNo = datas[xmlno];
   }
   return documentNo;
}
//---------------------------------------------
// URLの取得を行います。
//---------------------------------------------
function pdfurlSet(datas,directFlg) {
   var pdfUrl ='#';
   if(directFlg == 3){
	       pdfUrl = datas[xmlemaurl];
   } else if(directFlg == 4 ){
	       pdfUrl = datas[xmlepuburl];
   } else if(directFlg == 5){
	       pdfUrl = datas[xmldnpurl];
   } else {
	   if (!datas[xmlurl] && datas[xmlanchor]) {
	       pdfUrl = datas[xmlanchor];
	   } else if (!datas[xmlurl]) {
	   } else {
	       pdfUrl = datas[xmlurl];
	   }
	}
	if(DOC_DOMAIN_NAME != ""){
		pdfUrl = DOC_DOMAIN_NAME + pdfUrl;
	}
   return pdfUrl;
}
//---------------------------------------------
// バージョンの取得を行います。
//---------------------------------------------
function versionSet(datas,directFlg) {
   var version ='';
   if(directFlg == 3){
	  version = datas[xmlemaver];
   } else if(directFlg == 4 ){
	  version = datas[xmlepubver];
   } else if(directFlg == 5){
	  version = datas[xmldnpver];
   } else {
	  version = datas[xmlver];
   }
   return version;
}
//---------------------------------------------
// サイズの取得を行います。
//---------------------------------------------
function sizeSet(datas,directFlg) {
   var size ='';
   if(directFlg == 3){
	  size = datas[xmlemasize];
   } else if(directFlg == 4 ){
	  size = datas[xmlepubsize];
   } else if(directFlg == 5){
	  size = datas[xmldnpsize];
   } else {
	  size = datas[xmlpdfsize];
   }
   return size;
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// タイトルのHTMLイメージを作成します。
//---------------------------------------------
function titleSet(titleType, docNo, docVer, title, anchorTitle, anchorNewLine, directFlg, pdfUrl, directUrl, anchorUrl, no, pdfPage) {
   if (!titleType) { titleType = ""; }
   if (!docNo) { docNo = ""; }
   if (!docVer) { docVer = ""; }
   if (!title) { title = ""; }
   if (!anchorTitle) { anchorTitle = ""; }
   if (!anchorNewLine) { anchorNewLine = ""; }
   if (!directFlg) { directFlg = ""; }
   if (!pdfUrl) { pdfUrl = ""; }
   if (!directUrl) { directUrl = ""; }
   if (!anchorUrl) { anchorUrl = ""; }
   if (!no) { no = ""; }
   if (!pdfPage) { pdfPage = ""; }
   var titlehtm = new Array();
   // タイトルタイプ
   titlehtm.push("<tr><td class='docLinkSiryo'>");
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       var setPdfUrl = pdfUrl;
       if (pdfPage != null && pdfPage != "") {
           setPdfUrl = pdfUrl + "#page="+pdfPage;
       }
       titlehtm.push("<a href='" + setPdfUrl + "' title='" + docNo + "' target='_blank'>");
   } else if(directFlg == 2 && directUrl != "") {
       titlehtm.push("<a href='" + directUrl + "#" + no + "' title='" + docNo + "' target='_top'>");
   } else if(directFlg == 3 || directFlg == 4  || directFlg == 5) {
       titlehtm.push("<a href='" + pdfUrl +  "' title='" + docNo + "' target='_blank'>");
   } else {
   }
   if(titleType == "0" || titleType == "11"){
     //資料番号 + "-" + バージョン
       if (docVer != null && docVer != "") {
           titlehtm.push(docNo + "-" + docVer);
       } else {
           titlehtm.push(docNo);
       }
   }else if(titleType == "1"){
       //資料番号 + バージョン
       if (docVer != null && docVer != "") {
           titlehtm.push(docNo + docVer);
       } else {
           titlehtm.push(docNo);
       }
   }else if(titleType == "2"){
       //資料番号
       titlehtm.push(docNo);
   }else if(titleType == "3"){
       //資料名称
       var setTitle;
       if (anchorNewLine != null && anchorNewLine != "") {
         var buf1 = title.substring(0, anchorNewLine);
         var buf2 = title.substring(anchorNewLine);
         setTitle = buf1+"<br>"+buf2;
       } else {
         setTitle =title;
       }
       titlehtm.push(setTitle);
   }else if(titleType == "4"){
       //任意文字を設定
       if (anchorTitle == '') {
          titlehtm.push(docNo);
       } else {
          titlehtm.push(anchorTitle);
       }
   } else {
       return '';
   }
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       titlehtm.push("</a>");
   }else if(directFlg == 2 && directUrl != ""){
       titlehtm.push("</a>");
   } else if(directFlg == 3 || directFlg == 4  || directFlg == 5) {
       titlehtm.push("</a>");
   }else{
   }
   titlehtm.push("</td></tr>");
   return titlehtm.join("");
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// 画像のHTMLイメージを作成します。
//---------------------------------------------
function imageSet(docno, imageFile, directFlg, pdfUrl, directUrl, anchorUrl, imgWidth, imgHeight, no, imgBorder, pdfPage) {
   if (!docno) { docno = ""; }
   if (!imageFile) { imageFile = ""; }
   if (!directFlg) { directFlg = ""; }
   if (!pdfUrl) { pdfUrl = ""; }
   if (!directUrl) { directUrl = ""; }
   if (!anchorUrl) { anchorUrl = ""; }
   if (!imgWidth) { imgWidth = ""; }
   if (!imgHeight) { imgHeight = ""; }
   if (!no) { no = ""; }
   if (!imgBorder) { imgBorder = ""; }
   if (!pdfPage) { pdfPage = ""; }
   var imghtm = new Array();
   // 画像表示エリア
   imghtm.push("<tr><td class='docLinkImage'>");
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       if (no == anchorUrl ) {
           var setPdfUrl = pdfUrl;
           if (pdfPage != null && pdfPage != "") {
               setPdfUrl = pdfUrl + "#page="+pdfPage;
           }
           imghtm.push("<a href='" + setPdfUrl + "' title='" + docno + "' target='_blank'>");
       } else {
           imghtm.push("<a href='" + anchorUrl + "' title='" + docno + "' target='_top'>");
       }
   } else if(directFlg == 2  && directUrl != "") {
       imghtm.push("<a href='" + directUrl + "#" + no + "' title='" + docno + "' target='_top'>");
   } else if( directFlg == 3 || directFlg == 4 || directFlg == 5) {
       imghtm.push("<a href='" + pdfUrl + "' title='" + docno + "' target='_top'>");
   } else {
   }
   // 画像
   var setImgFile = "";
   if(IMG_DOMAIN_NAME != ""){
   		setImgFile = IMG_DOMAIN_NAME;
   }
   setImgFile += imageFile;
   if (imgBorder != null && imgBorder == "1") {
       imghtm.push("<img border='1' src='" + setImgFile + "'");
   } else {
       imghtm.push("<img src='" + setImgFile + "'");
   }

   // 横幅設定
   if (imgWidth != null && imgWidth != "") {
       imghtm.push(" width='"+ imgWidth  + "'");
   }
   // 縦幅設定
   if (imgHeight != null && imgHeight != "") {
       imghtm.push(" height='"+ imgHeight + "'");
   }
   imghtm.push(" >");

   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       imghtm.push("</a>");
   }else if(directFlg == 2  && directUrl != ""){
       imghtm.push("</a>");
   } else if( directFlg == 3 || directFlg == 4 || directFlg == 5) {
       imghtm.push("</a>");
   }else{
   }
   imghtm.push("</td></tr>");
   return imghtm.join("");
}

function pageLocation(directFlg, pdfurl, directurl, no, pdfPage) {
   if (directFlg == 1) {
       var setPdfUrl = pdfurl;
       if (pdfPage != null && pdfPage != "") {
           setPdfUrl = pdfurl + "#page="+pdfPage;
       }
       // PDFファイルのため別ウインドウで開く
       window.open(setPdfUrl);
       return true;
   } else if (directFlg == 2 && directurl != null && directurl != ""){
       // コンテンツのため同一ウインドウで開く
       top.location.href = directurl+ "#" + no;
       return true;
   } else if( directFlg == 3 || directFlg == 4 || directFlg == 5) {
       // PDFファイルのため別ウインドウで開く
       window.open(pdfurl);
       return true;
   }
   return false;
}

function getUrl(directFlg, pdfurl, directurl, pdfPage) {
   if (directFlg == 0 || directFlg == 1) {
       // PDFファイルのため別ウインドウで開く
       var setPdfUrl = pdfurl;
       if (pdfPage != null && pdfPage != "") {
           setPdfUrl = pdfurl + "#page="+pdfPage;
       }
       return setPdfUrl;
   } else if (directFlg == 2  && directurl != ""){
       // コンテンツのため同一ウインドウで開く
       return directurl;
   } else if (directFlg == 3  || directFlg == 4 || directFlg == 5){
       return pdfurl;
   }
   return "";
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// エラー時の出力を行います
//---------------------------------------------
function errOut(titleType, anchorTitle, docNo) {
   var outString = "";
   if (titleType != null && titleType == "4" && anchorTitle != null) {
      outString = anchorTitle;
   } else {
      if (docNo != null) {
         outString = docNo;
      }
   }
   return outString;
}

//---------------------------------------------
// ドキュメントリンク共通ロジック
// アップツールお知らせ用の出力を行います
//---------------------------------------------
function uptoolTopics(titleType, docNo, docVer, title, anchorTitle, anchorNewLine, directFlg, pdfUrl, directUrl, anchorUrl, no, pdfPage) {
   if (!titleType) { titleType = ""; }
   if (!docNo) { docNo = ""; }
   if (!docVer) { docVer = ""; }
   if (!title) { title = ""; }
   if (!anchorTitle) { anchorTitle = ""; }
   if (!anchorNewLine) { anchorNewLine = ""; }
   if (!directFlg) { directFlg = ""; }
   if (!pdfUrl) { pdfUrl = ""; }
   if (!directUrl) { directUrl = ""; }
   if (!anchorUrl) { anchorUrl = ""; }
   if (!no) { no = ""; }
   if (!pdfPage) { pdfPage = ""; }
   var titlehtm = new Array();
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       var setPdfUrl = pdfUrl;
       if (pdfPage != null && pdfPage != "") {
           setPdfUrl = pdfUrl + "#page="+pdfPage;
       }
       titlehtm.push("<a href='" + setPdfUrl + "' title='" + docNo + "' target='_blank'>");
   } else if(directFlg == 2 && directUrl != "") {
       titlehtm.push("<a href='" + directUrl + "#" + no + "' title='" + docNo + "' target='_top'>");
   } else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
       titlehtm.push("<a href='" + pdfUrl + "' title='" + docNo + "' target='_blank'>");
   } else {
   }
   //任意文字を設定
       //任意文字を設定
       if (anchorTitle == '') {
          titlehtm.push(docNo);
       } else {
          titlehtm.push(anchorTitle);
       }
//   titlehtm.push(anchorTitle);
   // ダイレクト表示・tラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       titlehtm.push("</a>");
   }else if(directFlg == 2 && directUrl != ""){
       titlehtm.push("</a>");
   } else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
       titlehtm.push("</a>");
   }else{
   }
   document.write(titlehtm.join(""));
}
function docNoOut(titleType, docNo, docVer, title, anchorTitle, anchorNewLine, directFlg, pdfUrl, directUrl, anchorUrl, no, pdfPage) {
   if (!titleType) { titleType = ""; }
   if (!docNo) { docNo = ""; }
   if (!docVer) { docVer = ""; }
   if (!title) { title = ""; }
   if (!anchorTitle) { anchorTitle = ""; }
   if (!anchorNewLine) { anchorNewLine = ""; }
   if (!directFlg) { directFlg = ""; }
   if (!pdfUrl) { pdfUrl = ""; }
   if (!directUrl) { directUrl = ""; }
   if (!anchorUrl) { anchorUrl = ""; }
   if (!no) { no = ""; }
   if (!pdfPage) { pdfPage = ""; }
   var titlehtm = new Array();
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       var setPdfUrl = pdfUrl;
       if (pdfPage != null && pdfPage != "") {
           setPdfUrl = pdfUrl + "#page="+pdfPage;
       }
       titlehtm.push("<a href='" + setPdfUrl + "' title='" + docNo + "'  target='_blank'>");
   } else if(directFlg == 2 && directUrl != "") {
       titlehtm.push("<a href='" + directUrl + "#" + no + "' title='" + docNo + "'  target='_top'>");
   } else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
       titlehtm.push("<a href='" + pdfUrl + "' title='" + docNo + "' target='_blank'>");
   } else {
   }
   //資料番号 + バージョン
   if (docVer != null && docVer != "") {
       titlehtm.push(docNo);
       titlehtm.push("-");
       titlehtm.push(docVer);
   } else {
       titlehtm.push(docNo);
   }
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       titlehtm.push("</a>");
   }else if(directFlg == 2 && directUrl != ""){
       titlehtm.push("</a>");
   } else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
       titlehtm.push("</a>");
   }else{
   }
   document.write(titlehtm.join(""));
}
function titleOut(titleType, docNo, docVer, title, anchorTitle, anchorNewLine, directFlg, pdfUrl, directUrl, anchorUrl, no, pdfPage) {
   if (!titleType) { titleType = ""; }
   if (!docNo) { docNo = ""; }
   if (!docVer) { docVer = ""; }
   if (!title) { title = ""; }
   if (!anchorTitle) { anchorTitle = ""; }
   if (!anchorNewLine) { anchorNewLine = ""; }
   if (!directFlg) { directFlg = ""; }
   if (!pdfUrl) { pdfUrl = ""; }
   if (!directUrl) { directUrl = ""; }
   if (!anchorUrl) { anchorUrl = ""; }
   if (!no) { no = ""; }
   if (!pdfPage) { pdfPage = ""; }
   var titlehtm = new Array();
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       var setPdfUrl = pdfUrl;
       if (pdfPage != null && pdfPage != "") {
           setPdfUrl = pdfUrl + "#page="+pdfPage;
       }
       titlehtm.push("<a href='" + setPdfUrl + "' title='" + docNo + "' target='_blank'>");
   } else if(directFlg == 2 && directUrl != "") {
       titlehtm.push("<a href='" + directUrl + "#" + no + "' title='" + docNo + "' target='_top'>");
   } else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
       titlehtm.push("<a href='" + pdfUrl + "' title='" + docNo + "' target='_blank'>");
   } else {
   }
   // タイトル
   titlehtm.push(title);
   // ダイレクト表示フラグ
   if(directFlg == 0){
   }else if(directFlg == 1){
       titlehtm.push("</a>");
   }else if(directFlg == 2 && directUrl != ""){
       titlehtm.push("</a>");
   } else if(directFlg == 3 || directFlg == 4 || directFlg == 5) {
       titlehtm.push("</a>");
   }else{
   }
   document.write(titlehtm.join(""));
}


/**
 * 応答データ構造体
 * 　documentno   :資料番号
 * 　pdfurl       :PDFのURL
 * 　pdfsize      :PDFサイズ（バイト）
 * 　pdfver       :PDFバージョン
 * 　documenttitle:資料のタイトル
 */
function documentdatas(documentno, pdfurl, pdfsize, pdfver, documenttitle){
	this.documentno=documentno;
	this.pdfurl=pdfurl;
	this.pdfsize=pdfsize;
	this.pdfver=pdfver;
	this.documenttitle=documenttitle;
}


function getDocumentDataList(docNos, kisyuId, type) {
	loadDataLinkinfo();
	var requesturl = new Array();
	requesturl.push(KISYU_NO);
	requesturl.push("=");
	requesturl.push(kisyuId);
	requesturl.push("&");
	requesturl.push(DOCUMENT_TYPE);
	requesturl.push("=");
	requesturl.push(type);
	var cnt =0;
	var roopflag=true;
	var ret = new Array();
	while(roopflag) {
		if (docNos.length) {
			
			for (i=cnt; i<docNos.length; i++) {
				cnt++;
				requesturl.push("&");
				requesturl.push(DOC_NO2);
				requesturl.push((i%10)+1);
				requesturl.push("=");
				requesturl.push(docNos[i]);
				if ( ((i%10)+1) == 10){
					break;
				}
			}
			if (cnt >= docNos.length) {
				roopflag=false;
			}
		} else if (docNos != null && docNos != "") {
			requesturl.push("&");
			requesturl.push(DOC_NO2);
			requesturl.push("1=");
			requesturl.push(docNos);
			roopflag=false;
		}
		// リクエスト送信
		var httpObj = createXmlHttp();
		if(httpObj == null){
			return;
		}
		var count = 0;
		var flg = false;
		for ( count=0; count<=KISYU_ARRAY.length; count++) {
			if (KISYU_ARRAY[count] == kisyuId) {
				flg = true;
				break;
			}
		}
		// 機種名での判定(機種IDでフラグがtrueになった場合はスキップ）
		if (!flg) {
		for ( count=0; count<=KISYU_NAME_ARRAY.length; count++) {
				if (KISYU_NAME_ARRAY[count] == kisyuId) {
					flg = true;
				break;
				}
			}
		}

	    // 通信をOPEN
	    if (flg) {
	        httpObj.open("POST", NEW_GET_DOCUMENT_LINK_INFO, false);
	    } else {
	        httpObj.open("POST", GET_DOCUMENT_LINK_INFO, false);
	    }
		httpObj.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		httpObj.send(requesturl.join(""));
		if(httpObj.readyState == 4){
			if(httpObj.status == 200){
				 // 返り値ＸＭＬの取得
				 xmlData  = httpObj.responseXML;
				 var errorcode = xmlData.getElementsByTagName("ErrorCD")[0].childNodes[0].nodeValue;
				 var workDocNumber = "";
				 if(errorcode == 0){
				 	if (xmlData.getElementsByTagName("DocAnchor") != null ) {
					 	var len = xmlData.getElementsByTagName("DocAnchor").length;
				 		for (j = 0; j < len; j++) {
						 	var ansXmlData = new Array();
							if(xmlData.getElementsByTagName("DocAnchor")[j] != null && xmlData.getElementsByTagName("DocAnchor")[j].childNodes[0] != null){
							 ansXmlData[xmlanchor] = xmlData.getElementsByTagName("DocAnchor")[j].childNodes[0].nodeValue;
							}
							
							if(xmlData.getElementsByTagName("DocNumber")[j] != null && xmlData.getElementsByTagName("DocNumber")[j].childNodes[0] != null){
							 ansXmlData[xmlno] = xmlData.getElementsByTagName("DocNumber")[j].childNodes[0].nodeValue;
							}
							
							if(xmlData.getElementsByTagName("DocDispNumber")[j] != null && xmlData.getElementsByTagName("DocDispNumber")[j].childNodes[0] != null){
							 ansXmlData[xmldispno] = xmlData.getElementsByTagName("DocDispNumber")[j].childNodes[0].nodeValue;
							}
							
							if(xmlData.getElementsByTagName("DocKisyuPath")[j] != null && xmlData.getElementsByTagName("DocKisyuPath")[j].childNodes[0] != null){
							 ansXmlData[xmlkisyupath] = xmlData.getElementsByTagName("DocKisyuPath")[j].childNodes[0].nodeValue;
							}
							
							if(xmlData.getElementsByTagName("DocPdfURL")[j] != null && xmlData.getElementsByTagName("DocPdfURL")[j].childNodes[0] != null){
								ansXmlData[xmlurl] = xmlData.getElementsByTagName("DocPdfURL")[j].childNodes[0].nodeValue;
							}
							
							if (xmlData.getElementsByTagName("DocPdfVer")[j] != null && xmlData.getElementsByTagName("DocPdfVer")[j].childNodes[0] != null) {
								ansXmlData[xmlver] = xmlData.getElementsByTagName("DocPdfVer")[j].childNodes[0].nodeValue;
							}
							
							if(xmlData.getElementsByTagName("DocTitle")[j] != null && xmlData.getElementsByTagName("DocTitle")[j].childNodes[0] != null){
							 ansXmlData[xmltitle] = xmlData.getElementsByTagName("DocTitle")[j].childNodes[0].nodeValue;
							}
							
							if(xmlData.getElementsByTagName("DocImgKisyuURL")[j] != null && xmlData.getElementsByTagName("DocImgKisyuURL")[j].childNodes[0] != null){
							 ansXmlData[xmlimgKisyu] = xmlData.getElementsByTagName("DocImgKisyuURL")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocImgCatalogURL")[j] != null && xmlData.getElementsByTagName("DocImgCatalogURL")[j].childNodes[0] != null){
							 ansXmlData[xmlimgCatalog] = xmlData.getElementsByTagName("DocImgCatalogURL")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocImgOptionURL")[j].childNodes[0] != null && xmlData.getElementsByTagName("DocImgOptionURL")[j].childNodes[0] != null){
							 ansXmlData[xmlimgOption] = xmlData.getElementsByTagName("DocImgOptionURL")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocDate")[j] != null && xmlData.getElementsByTagName("DocDate")[j].childNodes[0] != null){
							 var date = xmlData.getElementsByTagName("DocDate")[j].childNodes[0].nodeValue;
							 var buff = date.split("/");
							 //YYYY/MMへ変更
							 //ansXmlData[xmldate] = buff[0] + "年" + buff[1] + "月";
							 ansXmlData[xmldate] = buff[0] + "-" + buff[1];
							}
							if(xmlData.getElementsByTagName("DocStatus")[j] != null && xmlData.getElementsByTagName("DocStatus")[j].childNodes[0] != null){
							 ansXmlData[xmlshiryoStatus] = xmlData.getElementsByTagName("DocStatus")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocRegistType")[j] != null && xmlData.getElementsByTagName("DocRegistType")[j].childNodes[0] != null){
							 ansXmlData[xmlshiryoDel] = xmlData.getElementsByTagName("DocRegistType")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocBunruiL")[j] != null && xmlData.getElementsByTagName("DocBunruiL")[j].childNodes[0] != null){
							 ansXmlData[xmlbunruil] = xmlData.getElementsByTagName("DocBunruiL")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocBunruiS")[j] != null && xmlData.getElementsByTagName("DocBunruiS")[j].childNodes[0] != null){
							 ansXmlData[xmlbunruis] = xmlData.getElementsByTagName("DocBunruiS")[j].childNodes[0].nodeValue;
							}
							if(xmlData.getElementsByTagName("DocPdfSize")[j] != null && xmlData.getElementsByTagName("DocPdfSize")[j].childNodes[0] != null){
							 ansXmlData[xmlpdfsize] = xmlData.getElementsByTagName("DocPdfSize")[j].childNodes[0].nodeValue;
							}
				 			if (workDocNumber != ansXmlData[xmlno]) {
				 				ret.push( new documentdatas(ansXmlData[xmlno], ansXmlData[xmlurl], ansXmlData[xmlpdfsize], ansXmlData[xmlver],  ansXmlData[xmltitle]) );
				 				workDocNumber = ansXmlData[xmlno];
				 			}
						   // ドキュメントDB改善　フェーズ2対応
						   // E-manual
						   if(xmlData.getElementsByTagName("DocEmaURL")[j] != null && xmlData.getElementsByTagName("DocEmaURL")[j].childNodes[0] != null){
						       ansXmlData[xmlemaurl] = xmlData.getElementsByTagName("DocEmaURL")[j].childNodes[0].nodeValue;
						   }
						   
						   if (xmlData.getElementsByTagName("DocEmaVer")[j] != null && xmlData.getElementsByTagName("DocEmaVer")[j].childNodes[0] != null) {
						       ansXmlData[xmlemaver] = xmlData.getElementsByTagName("DocEmaVer")[j].childNodes[0].nodeValue;
						   }

						   if(xmlData.getElementsByTagName("DocEmaSize")[j] != null && xmlData.getElementsByTagName("DocEmaSize")[j].childNodes[0] != null){
						     ansXmlData[xmlemasize] = xmlData.getElementsByTagName("DocEmaSize")[j].childNodes[0].nodeValue;
						   }

							// epub
						   if(xmlData.getElementsByTagName("DocEpubURL")[j] != null && xmlData.getElementsByTagName("DocEpubURL")[j].childNodes[0] != null){
						       ansXmlData[xmlepuburl] = xmlData.getElementsByTagName("DocEpubURL")[j].childNodes[0].nodeValue;
						   }
						   
						   if (xmlData.getElementsByTagName("DocEpubVer")[j] != null && xmlData.getElementsByTagName("DocEpubVer")[j].childNodes[0] != null) {
						       ansXmlData[xmlepubver] = xmlData.getElementsByTagName("DocEpubVer")[j].childNodes[0].nodeValue;
						   }

						   if(xmlData.getElementsByTagName("DocEpubSize")[j] != null && xmlData.getElementsByTagName("DocEpubSize")[j].childNodes[0] != null){
						     ansXmlData[xmlepubsize] = xmlData.getElementsByTagName("DocEpubSize")[j].childNodes[0].nodeValue;
						   }

							//dnp
						   if(xmlData.getElementsByTagName("DocDnpURL")[j] != null && xmlData.getElementsByTagName("DocDnpURL")[j].childNodes[0] != null){
						       ansXmlData[xmldnpurl] = xmlData.getElementsByTagName("DocDnpURL")[j].childNodes[0].nodeValue;
						   }
						   
						   if (xmlData.getElementsByTagName("DocDnpVer")[j] != null && xmlData.getElementsByTagName("DocDnpVer")[j].childNodes[0] != null) {
						       ansXmlData[xmldnpver] = xmlData.getElementsByTagName("DocDnpVer")[j].childNodes[0].nodeValue;
						   }
						   
						   if(xmlData.getElementsByTagName("DocDnpSize")[j] != null && xmlData.getElementsByTagName("DocDnpSize")[j].childNodes[0] != null){
						     ansXmlData[xmldnpsize] = xmlData.getElementsByTagName("DocDnpSize")[j].childNodes[0].nodeValue;
						   }
				 		}
				 	}
				 }
			}
		}
		if (roopflag) {
			requesturl = new Array();
			requesturl.push(KISYU_NO);
			requesturl.push("=");
			requesturl.push(kisyuId);
			requesturl.push("&");
			requesturl.push(DOCUMENT_TYPE);
			requesturl.push("=");
			requesturl.push(type);
		}
	}
	return ret;
}
