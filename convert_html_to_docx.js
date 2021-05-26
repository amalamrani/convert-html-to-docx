/* Parse document HTML to DOCX 
*  Author: Amal Amrani
*  AUthor URL: http://amrani.es
*/



//  PREVIOUSLY WE CONVERT ALL DOCUMENT IMAGES TO BASE64



function funcCallback(){
    
}

/**
* recursive function to convert all document images to Base64
*/

function convertAllUrlImagesToBase64(docu){

    var allImgDoc = docu.getElementsByTagName("IMG");
    var allImg = []; 

    var i=0;
    for( var j=0; j< allImgDoc.length; j++ ){ 

        var url = allImgDoc[j].src;

        if ( url.startsWith('http:') && url.includes(window.location.hostname) && url.match(/.(jpg|jpeg|png|gif)/gi) )
            allImg.push(allImgDoc[j]);
    }    

    if ( allImg.length )      recursiva_img( allImg, i, funcCallback);   
}


/**
* function to be executed when recursiva_img finished
*/

function ConversionCompleted(){

    // Call function to generat docx
    var myDocx = generateDocx(docu);
   
   // included in generateDocx function!
    /*saveAs(myDocx, "example.docx");*/

}

/**
* Function to draw document image in canvas to convert it from url to Base64
*/

function recursiva_img(allImg,i, funcCallback){    

    var url = allImg[i].src;    
    var res = null;

        
                // createCanvas 
                var img = new Image();
                img.crossOrigin = 'Anonymous';
                img.src = url;
                
                img.onload = function(){

                    var canvas = document.createElement('CANVAS');
                    ctx = canvas.getContext('2d');
                    canvas.height = img.height;
                    canvas.width = img.width;
                    ctx.drawImage(img, 0, 0);
                    
                    var dataURL = canvas.toDataURL('image/png');
                    
                    allImg[i].src = dataURL;
                    allImg[i].width = canvas.width;
                    allImg[i].height = canvas.height;
                    
                    
                    canvas = null
                    i++; 
                    
                    if( i < allImg.length ){
                     
                    return recursiva_img(allImg,i);
                    }    
                    else{
                    
                        ConversionCompleted();                    
                    } 
                };
}
//// END CONVERT IMAGES TO BASE64


/**
* function parse HTML to Docx document. Need jszip.js and FileSaver.js
* @param: html string
* 
*/

function generateDocx(doc){ 
	
  	//   crear los nodos del xml
	var nodeParent = doc.getElementsByTagName("body"); 
	var xmlDoc = document.implementation.createDocument(null, "mywordXML");

	var padreXML = xmlDoc.childNodes[0]; 

	var zip = new JSZip();


	// VARIABLES FUNCION PRINCIPAL
	var numImg = 1;
	var numLink = 1;
	var relsDocumentXML = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
					  '<Relationship Id="rId0" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>'; // styles.xml del document

	var countList = 0;
	var itemsList = 0;
	var ultimoItemAnidado = false;

	var numberingString= '';
	var numIdString='';

	var cols = 0;

	var tabla = false;
	var orientationRTL = false;
	
	this.checkIfInsertTag=function (newEle){

		

		if(newEle.nodeName === 'w:Noinsert')  return false; 
		
		else return true;

		
	}	
	

	this.createNodeBlockquote=function (node,xmlDoc){

		var newEle = xmlDoc.createElement("w:p");
				var wpPr = xmlDoc.createElement("w:pPr");

					var wstyle= xmlDoc.createElement("w:pStyle");  wstyle.setAttribute('w:val','blockQuote');
				wpPr.appendChild(wstyle);
				
			newEle.appendChild(wpPr);	

			
	return newEle;
	}		

	/**
		*  función que comprueba si existe consanguinidad en la rama con una lista para insertar la rama en el padre = w:body
		*  comprobación de cada nodo antes de sacar su equivalente en el docx
		*  ul, ol, li, y cualquier otro elemento que tenga consanguinidad con una lista
		*  Se entiende que un elemento tiene consanguinidad con lista si su padre no es body y además alguno de sus parientes (hermanos o descendientes como hijos, sobrinos, nietos o hijos de sobrinos, bisnietos, etc)
		*  @param node el nodo originar a comprobar si tiene consanguinidad con una lista
		*  @return boolen true or false
		* 
		*/
		this.consanguinidadLista=function (node){

			if(node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD')  return false;

			// SI EL ELEMENTOS ES UNO DE ESTOS DIRECTAMENTE DEVOLVER TRUE 			
			if( node.nodeName === 'UL' || node.nodeName === 'OL' || node.nodeName === 'LI'  )   return true;
			
			// SI EL PADRE NO ES BODY Y TIENEN CONSANGUINIDAD = > TRUE
			else{ 
				
				var padre = node.parentNode;

				// PREGUNTA A HIJOS Y DESCENDIENTES
				for(var hijo=0; hijo<node.childNodes.length;hijo++)
					if(this.consanguinidadLista(node.childNodes[hijo]))  return true;					
			
				// PREGUNTA A HERMANOS Y DESCENDIENTES
				for( var hermano=0; hermano< padre.childNodes.length; hermano++ ){
						// si el hermano no soy yo preguntar
						if ( padre.childNodes[hermano] != node){

							var nodHermano = padre.childNodes[hermano];
							for(var sobrino = 0; sobrino < nodHermano.childNodes.length; sobrino++)
								if(this.consanguinidadLista(nodHermano.childNodes[sobrino])) return true;
															
						}

				}
												
			}

		return false;
		}

		
	// FUNCION RECURSIVE TO CREATE DOCX XML NODE 
	this.createNodeXML = function(node,xmlDoc,padreXML){

	
		var newEle = parseHTMLtoDocx(node,xmlDoc,padreXML);
		
		// IMPORTANTE SACAR AQUI EL ANCESTRO Y NO ANTES. PARA ASEGURARNOS DE QUE SE HA CREADO YA EL ELEMNTO BODY
		var ancestro = xmlDoc.getElementsByTagName('w:body')[0];  
		
		var insertarTag = this.checkIfInsertTag(newEle);
		
		if(insertarTag){

			// IF TABLA
			if( newEle.nodeName === 'w:tbl' ){  

				// insertar en ancestro
				ancestro.appendChild(newEle);
				tabla = newEle;

			}
						
			else if( newEle.nodeName === 'w:tr' ){			

				 tabla.appendChild(newEle);

			}
			
			//  SI SE TRATA DE UNA LISTA, INSERTAR EN ANCESTRO = W:BODY			
			else if( this.consanguinidadLista(node)){ 
			    				
					if( newEle.nodeName === 'w:r'){ 
					
				        if(ancestro.lastElementChild.nodeName === 'w:p') ancestro.lastElementChild.appendChild(newEle);
				        
				        else{
				        	
				        	var p = xmlDoc.createElement("w:p");   p.appendChild(newEle);  ancestro.appendChild(p);  
				    	
				    	}
				    }
					else ancestro.appendChild(newEle);				

			}

			else{  
			
		        padreXML.appendChild(newEle);  
			}

		}
	    	  
  	    if (newEle) padreXML = newEle;  
	    	
		for(var hijo =0; hijo < node.childNodes.length; hijo++)
			 
			createNodeXML(node.childNodes[hijo],xmlDoc,padreXML);

	}

		this.createNodeTD=function (node,xmlDoc,cols){ 

			var widthPercent = 100/cols;

			var newEle = xmlDoc.createElement("w:tc");
				var wtcPr = xmlDoc.createElement("w:tcPr");
					var valign = xmlDoc.createElement("w:vAlign");  valign.setAttribute('w:val','bottom');
					var wtblBorders = xmlDoc.createElement("w:tblBorders");
						var wtop = xmlDoc.createElement("w:top"); wtop.setAttribute('w:val','single'); wtop.setAttribute('w:sz','10'); wtop.setAttribute('w:space','0'); wtop.setAttribute('w:color','000000');			
						var wstart = xmlDoc.createElement("w:start"); wstart.setAttribute('w:val','single'); wstart.setAttribute('w:sz','10'); wstart.setAttribute('w:space','0'); wstart.setAttribute('w:color','000000');			
						var wbottom = xmlDoc.createElement("w:bottom"); wbottom.setAttribute('w:val','single'); wbottom.setAttribute('w:sz','10'); wbottom.setAttribute('w:space','0'); wbottom.setAttribute('w:color','000000');			
						var wend = xmlDoc.createElement("w:end"); wend.setAttribute('w:val','single'); wend.setAttribute('w:sz','10'); wend.setAttribute('w:space','0'); wend.setAttribute('w:color','000000');
					wtblBorders.appendChild(wtop);	
					wtblBorders.appendChild(wstart);	
					wtblBorders.appendChild(wbottom);	
					wtblBorders.appendChild(wend);	
					//  color thead si existe
					var granFather = (node.parentNode).parentNode;
					if ( granFather.nodeName === 'THEAD'){
					var wshd = xmlDoc.createElement("w:shd"); wshd.setAttribute('w:val','clear'); wshd.setAttribute('w:fill','EEEEEE');
					wtcPr.appendChild(wshd); 
					}
					

					var wtcW = xmlDoc.createElement("w:tcW");  wtcW.setAttribute('w:type','pct');   wtcW.setAttribute('w:w',widthPercent+'%');
				
				
				wtcPr.appendChild(wtblBorders);
				wtcPr.appendChild(wtcW);
				wtcPr.appendChild(valign);

			newEle.appendChild(wtcPr);

			
			// SI TD VACÍO 
			if ( node.childNodes.length == 0 ) {
				var nodep =  xmlDoc.createElement("w:p");
				
				var noder =  xmlDoc.createElement("w:r");
				
				var nodetext = xmlDoc.createElement("w:t");
					var text = xmlDoc.createTextNode(" ");
					nodetext.appendChild(text);
					noder.appendChild(nodetext);
				nodep.appendChild(noder);
			newEle.appendChild(nodep);
			}

		return newEle;
		}	
	
		this.createNodeTR=function (node,xmlDoc){
			
			return xmlDoc.createElement("w:tr");
		}

		this.createTableNode=function (node,xmlDoc){

			var newEle = xmlDoc.createElement("w:tbl"); 
				var wtblPr = xmlDoc.createElement("w:tblPr");
					var wtblStyle = xmlDoc.createElement("w:tblStyle");  wtblStyle.setAttribute('w:val','TableGrid');
					var wtblW = xmlDoc.createElement("w:tblW");	wtblW.setAttribute('w:w','5000'); wtblW.setAttribute('w:type','pct');	
					wtblPr.appendChild(wtblStyle);
					wtblPr.appendChild(wtblW);

					var wtblBorders = xmlDoc.createElement("w:tblBorders");
						var wtop = xmlDoc.createElement("w:top"); wtop.setAttribute('w:val','single'); wtop.setAttribute('w:sz','10'); wtop.setAttribute('w:space','0'); wtop.setAttribute('w:color','000000');
						var wstart = xmlDoc.createElement("w:start"); wstart.setAttribute('w:val','single'); wstart.setAttribute('w:sz','10'); wstart.setAttribute('w:space','0'); wstart.setAttribute('w:color','000000');
						var wbottom = xmlDoc.createElement("w:bottom"); wbottom.setAttribute('w:val','single'); wbottom.setAttribute('w:sz','10'); wbottom.setAttribute('w:space','0'); wbottom.setAttribute('w:color','000000');
						var wend = xmlDoc.createElement("w:end"); wend.setAttribute('w:val','single'); wend.setAttribute('w:sz','10'); wend.setAttribute('w:space','0'); wend.setAttribute('w:color','000000');

						var windideH =  xmlDoc.createElement("w:insideH");  windideH.setAttribute('w:val','single'); windideH.setAttribute('w:sz','5'); windideH.setAttribute('w:space','0'); windideH.setAttribute('w:color','000000');
						var windideV =  xmlDoc.createElement("w:insideV");  windideV.setAttribute('w:val','single'); windideV.setAttribute('w:sz','5'); windideV.setAttribute('w:space','0'); windideV.setAttribute('w:color','000000');


						wtblBorders.appendChild(wtop);
						wtblBorders.appendChild(wstart);
						wtblBorders.appendChild(wbottom);
						wtblBorders.appendChild(wend);

						wtblBorders.appendChild(windideH);
						wtblBorders.appendChild(windideV);

					wtblPr.appendChild(wtblBorders);

			newEle.appendChild(wtblPr);

		return newEle;

		}

		this.createHeading=function (node,xmlDoc,pos){


			if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD'){
			
					var newEle = xmlDoc.createElement("w:p");
					
				
					 var newEle2 = xmlDoc.createElement("w:pPr");

					var newEle3 = xmlDoc.createElement("w:pStyle");  newEle3.setAttribute('w:val', 'Heading'+pos);

					newEle.appendChild(newEle2);
					newEle2.appendChild(newEle3);

					newEle = checkRTL(newEle,xmlDoc,node);
			}
			else{

				var newEle = xmlDoc.createElement("w:r");
				

			}

					
		return newEle;
		}
		this.checkRTL = function(newEle,xmlDoc,node){ 

				if( node.attributes && node.attributes[0] ){
			       		
			       		var direction = node.attributes[0].nodeValue.replace(' ',''); 
			       	
			       		direction = direction.replace('direction:','');
			       		
			       		 var res = direction.substring(0, 3); 
			       		if ( res === 'rtl'){
			       			orientationRTL = true;
			       			//AÑADIMOS AQUI RTL 
 							if(newEle.childNodes && newEle.childNodes[0] && newEle.childNodes[0].nodeName === 'w:pPr'){ 
 								
 								var elepPr = newEle.childNodes[0]; 
 						    }
			       			else var elepPr = xmlDoc.createElement("w:pPr");
			       				var bidi = xmlDoc.createElement("w:bidi");  bidi.setAttribute('w:val','1');
			       				elepPr.appendChild(bidi);
			       			
			       			
						newEle.appendChild(elepPr);
			       		}
			       		else orientationRTL = false;
			        }
			       
		return newEle;
		}

		this.createNodeStrong=function (node, xmlDoc){
			var newEle;

			var newEleR = xmlDoc.createElement("w:r");  
			var bold = xmlDoc.createElement("w:rPr");
			var propertyBold = xmlDoc.createElement("w:b");
			newEleR.appendChild(bold);
			bold.appendChild(propertyBold);

			if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD' ){

				newEle = xmlDoc.createElement("w:p"); 
				newEle.appendChild(newEleR);
			}
			else newEle = newEleR;

		return newEle;			

		}

		this.createNodeEM=function (node, xmlDoc){

			var newEle;

			var newEleR = xmlDoc.createElement("w:r");  
			var nodeProperty = xmlDoc.createElement("w:rPr");
			var property = xmlDoc.createElement("w:i");
			newEleR.appendChild(nodeProperty);
			nodeProperty.appendChild(property);

			if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD' ){

				newEle = xmlDoc.createElement("w:p"); 
				newEle.appendChild(newEleR);
			}
			else newEle = newEleR;

		return newEle;
		}

		this.createNodeBR = function(node, xmlDoc){

			var newEle;

			var newEleR = xmlDoc.createElement("w:r");
				var eleBR = xmlDoc.createElement("w:br");
			newEleR.appendChild(eleBR);


			if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD' ){

				newEle = xmlDoc.createElement("w:p"); 
				newEle.appendChild(newEleR);
			}
			else newEle = newEleR;


		return newEle;
		}

		this.createHiperlinkNode=function (node,xmlDoc,numLink){ 
			// SI CONTIENE UNA IMAGEN
			for( var j=0;j<node.childNodes.length;j++){
				if(node.childNodes[j].nodeName === 'IMG'){

					var nR = xmlDoc.createElement("w:r");
					var nT = xmlDoc.createElement("w:t"); nT.setAttribute('xml:space','preserve');
					var t  = xmlDoc.createTextNode("link format not available");
					nR.appendChild(nT);
					nT.appendChild(t);

					if(node.parentNode.nodeName === 'BODY' ){ var p = xmlDoc.createElement("w:p"); p.appendChild(nR); return p; }
					else return nR;
				}
			}


			var hiperEle = xmlDoc.createElement("w:hyperlink");  
		    if (node.href ){  // EXTERNAL LINK
		    	hiperEle.setAttribute('r:id', 'link'+numLink);  
		    
		    }
		

			if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD' ){

						var newEle = xmlDoc.createElement("w:p"); 

						newEle.appendChild(hiperEle);

			}
			else	var newEle = hiperEle;

		return newEle;
		}

		this.createNodeLi=function (node,xmlDoc,countList,itemsList){

			
			var father = node.parentNode;

			var newEle = xmlDoc.createElement("w:p"); 
			var pr =  xmlDoc.createElement("w:pPr");
			var rstyle=  xmlDoc.createElement("w:pStyle");  rstyle.setAttribute('w:val','ListParagraph'); //('mystyleList');   ListParagraph 
			
			pr.appendChild(rstyle);
			newEle.appendChild(pr);

				var numpr = xmlDoc.createElement("w:numPr");
					var wilvl = xmlDoc.createElement("w:ilvl");   wilvl.setAttribute('w:val',itemsList);
					var numId = xmlDoc.createElement("w:numId");    numId.setAttribute('w:val',countList); 
						
					numpr.appendChild(wilvl);
					numpr.appendChild(numId);				

				pr.appendChild(numpr);


		return newEle;
		}		

		/**
		*  function to create a xml document node paragraph or node run
		*  @param: node html node to parse
		*  @return a new document.xml node
		*/
		this.createNodeParagraphOrRun=function (node,xmlDoc){
			var tx = node.data;

			    	// si texto undefined. asignar caracter blanco
			    	if (!tx)  tx = '';   
			    	
			       // SI EL TAG NO RECONOCIBLE Y PADRE = BODY. SE TRADUCE EN UN PÁRRAFO
			       if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD' ){

			       	var newEle = xmlDoc.createElement("w:p");

			       	newEle = checkRTL(newEle,xmlDoc,node);			       
					
			       }
			       else{  
			       	
				       	var newEle = xmlDoc.createElement("w:r");
				      
			       }


		return newEle;
		}		

		/**
		*  function to create text node
		*  @param node html node to convert into docx node
		*  @param xmlDoc xml document
		*  @return docx node
		*/
		this.createMyTextNode=function (node,xmlDoc){
			
			var tx = node.data;   

			if (!tx) tx = ''; // IF TX UNDEFINED  !!!!


			if( node.parentNode.nodeName === 'TR' || node.parentNode.nodeName === 'TABLE'  )  return xmlDoc.createElement("w:Noinsert"); 
			else if ( node.parentNode.nodeName === 'BODY' &&  ( !node.data.trim() ) )  return xmlDoc.createElement("w:Noinsert"); 
			  			 			
				var newEleR = xmlDoc.createElement("w:r"); 

				
			        if( orientationRTL ){

			        	var elerPr = xmlDoc.createElement("w:rPr");
			       					var rtl = xmlDoc.createElement("w:rtl");  rtl.setAttribute('w:val','1');
			       					elerPr.appendChild(rtl);			       				
			       			
					newEleR.appendChild(elerPr);
			        }

				if( node.parentNode.nodeName === 'A' && node.parentNode.href){
					var nodeStyleLink = xmlDoc.createElement("w:rStyle");    nodeStyleLink.setAttribute("w:val","Hyperlink");
					newEleR.appendChild(nodeStyleLink);
				}

		    	var nodetext = xmlDoc.createElement("w:t");  nodetext.setAttribute('xml:space', 'preserve');
		    	var texto = xmlDoc.createTextNode(tx);

		    	
		    	newEleR.appendChild(nodetext);
		    	nodetext.appendChild(texto);

		    if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD'){
				
				var newEle = xmlDoc.createElement("w:p");
				newEle.appendChild(newEleR);
			}			
			else  var newEle = newEleR;

			

		return newEle;
		}	

		this.createAstractNumListDecimal = function(numberingString,countList){
			numberingString= `<w:abstractNum w:abstractNumId="`+countList+`" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:multiLevelType w:val="hybridMultilevel" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
		<w:lvl w:ilvl="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" />
			<w:numFmt w:val="decimal" />
			<w:lvlText w:val="%1." />
			<w:lvlJc w:val="left" />
			<w:pPr><w:ind w:left="720" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%2." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="1440" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="2" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%3." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="2160" w:hanging="180" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="3" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%4." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="2880" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="4" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%5." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="3600" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="5" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%6." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="4320" w:hanging="180" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="6" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%7." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5040" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="7" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%8." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5760" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="8" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%9." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="6480" w:hanging="180" /></w:pPr>
		</w:lvl>
</w:abstractNum>`+numberingString;


			
		return numberingString;
		}	

		this.createAstractNumListBullet = function(numberingString,countList){
			numberingString= `<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="`+countList+`">
	<w:multiLevelType xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="hybridMultilevel"/>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="0">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val=""/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="720" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="1">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val="o"/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="1440" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="2">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val=""/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="2160" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="3">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="2880" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="4">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val="o"/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="3600" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="5">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="4320" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="6">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="5040" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="7">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val="o"/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="5760" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="8">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="6480" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
	</w:rPr>
	</w:lvl>
</w:abstractNum>`+numberingString;


			
		return numberingString;
		}	
		
		this.createNumIdList = function(numIdString, countList){

			numIdString= '<w:num w:numId="'+countList+'"><w:abstractNumId w:val="'+countList+'"/></w:num>'+numIdString;
			
		return numIdString;

		}

		/**
		* function para redimensionar la imagen si es necesario y excede el tamaño de página
		* devuelve las dimensiones de la imagen en EMUs para pintarla posteriormente en el docx
		*
		*  hacer la conversión corrcetamente teniendo en cuanta DPI
		*	
		*	1 inch (pulgada)   = 914400 EMUs
		*
		*	1cm = 360000 EMUs
		*	Teniendo en cuenta que la página mide 21 x 29.7 cm	
		*
		*
		*/
		this.escalarIMG=function (imgwidth, imgheight){

			var dimensionIMG = {'width': 0, 'height': 0}

			// tomando  96px por pulgada . si se quiere la imagen para impresion en papel probar dividir entre 300 x pulgada 
			var width_inch = imgwidth / 96;  
		  	var height_inch = imgheight / 96; 

		  	var width_emu = width_inch * 914400;
		  	var height_emu = height_inch * 914400;

		  	var pgSzW =  (16 * 360000); // // ancho de página en EMUs
		  	var pgSzH = (24.7 * 360000); // // alto de página en EMUs


		  
		  	if (width_emu > pgSzW){  
		  	var originalW = width_emu;
		  	width_emu = pgSzW;  
		  	height_emu = Math.floor (width_emu * height_emu / originalW);

		  	}

		  	if (height_emu > pgSzH){ 

		  	var originalH = height_emu;
		  	height_emu = pgSzH;  
		  	width_emu = Math.floor (height_emu * width_emu / originalH);

		  	
		  	}

		  	dimensionIMG.width =  Math.floor(width_emu);
		  	dimensionIMG.height =  Math.floor(height_emu);
		  			    

		return dimensionIMG;
		}

		/**
		*  Función que devuelve nodo vacío porque no reconoce el formato de imagen
		*  @param xmlDoc xml docx document 
		*  @return node element with text = Not available image format
		*/
		this.nodeVoid=function (xmlDoc){

			var newEle = xmlDoc.createElement("w:r"); 
				var nodetext = xmlDoc.createElement("w:t");  nodetext.setAttribute('xml:space', 'preserve');
			    	var texto = xmlDoc.createTextNode('IMAGE FORMAT NOT AVAILABLE!');

			    	
			    	newEle.appendChild(nodetext);
			    	nodetext.appendChild(texto);


				return newEle;

		}

		/** 
		*  función que crea el nodo picture a insertar en el document.xml de docx
		*  @param: node  el nodo IMG origen a parsear
		*  
		*/

		this.createDrawingNodeIMG=function (node,dataImg,xmlDoc,numImg){
			
				
			var format = '.png';
			var nameFile = 'image'+numImg+format;  
				
			var relashionImg = 'rId'+numImg;


			// CREO NODO IMAGEN EN EL document PARA PODER VER SU ALTO Y ANCHO
			var img = document.createElement('img');
		  	img.src= dataImg;  
		  	img.width = node.width;
		  	img.height = node.height;


			// ESCALAMOS IMAGEN
		  	var dimensionImg = this.escalarIMG(img.width, img.height);


			// CREAR NODO DRAWING PICTURE EN DOCX

			var newEleImage = xmlDoc.createElement("w:r"); 

				var drawEle = xmlDoc.createElement("w:drawing");  
					var wpinline = xmlDoc.createElement("wp:inline");          wpinline.setAttribute('distR', '0');  wpinline.setAttribute('distL', '0');  wpinline.setAttribute('distB', '0');  wpinline.setAttribute('distT', '0');
						var wpextent = xmlDoc.createElement("wp:extent");       wpextent.setAttribute('cy', dimensionImg.height); wpextent.setAttribute('cx', dimensionImg.width);
						var wpeffectExtent = xmlDoc.createElement("wp:effectExtent");       wpeffectExtent.setAttribute('b', '0');      wpeffectExtent.setAttribute('r', '0');  wpeffectExtent.setAttribute('t', '0'); wpeffectExtent.setAttribute('l', '0');
						var wpdocPr = xmlDoc.createElement("wp:docPr");    wpdocPr.setAttribute('name', nameFile ); wpdocPr.setAttribute('id', numImg ); 

						var wpcNvGraphicFramePr = xmlDoc.createElement("wp:cNvGraphicFramePr"); 
						var childcNvGraphicFramePr = xmlDoc.createElement("a:graphicFrameLocks");     childcNvGraphicFramePr.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main' ); childcNvGraphicFramePr.setAttribute('noChangeAspect', '1' );
					   	
					   	wpcNvGraphicFramePr.appendChild(childcNvGraphicFramePr);

					   	var agraphic = xmlDoc.createElement("a:graphic"); agraphic.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main' );
					   		var agraphicdata = xmlDoc.createElement("a:graphicData"); agraphicdata.setAttribute('uri','http://schemas.openxmlformats.org/drawingml/2006/picture');
								var pic = xmlDoc.createElement("pic:pic"); pic.setAttribute('xmlns:pic','http://schemas.openxmlformats.org/drawingml/2006/picture');
									var picnvPicPr = xmlDoc.createElement("pic:nvPicPr"); 
										var piccNvPr = xmlDoc.createElement("pic:cNvPr");   piccNvPr.setAttribute('name', nameFile); piccNvPr.setAttribute('id', numImg);
										var piccNvPicPr = xmlDoc.createElement("pic:cNvPicPr");
										picnvPicPr.appendChild(piccNvPr);
										picnvPicPr.appendChild(piccNvPicPr);

									var picblipFill = xmlDoc.createElement("pic:blipFill");
										var ablip = xmlDoc.createElement("a:blip");     ablip.setAttribute('cstate', 'print' ); ablip.setAttribute('r:embed', relashionImg );
										var astretch = xmlDoc.createElement("a:stretch");
											var afillRect =  xmlDoc.createElement("a:fillRect");
											astretch.appendChild(afillRect);

										picblipFill.appendChild(ablip);
										picblipFill.appendChild(astretch);
										

									var picspPr = xmlDoc.createElement("pic:spPr");
										var axfrm = xmlDoc.createElement("a:xfrm");
											var aoff = xmlDoc.createElement("a:off");   aoff.setAttribute('y','0'); aoff.setAttribute('x','0');
											var aext = xmlDoc.createElement("a:ext");   aext.setAttribute('cy', dimensionImg.height); aext.setAttribute('cx', dimensionImg.width);
											axfrm.appendChild(aoff);
											axfrm.appendChild(aext);
											

										var aprstGeom = xmlDoc.createElement("a:prstGeom");  aprstGeom.setAttribute('prst','rect');
											var aavLst = xmlDoc.createElement("a:avLst");
											aprstGeom.appendChild(aavLst);

										picspPr.appendChild(axfrm);
										picspPr.appendChild(aprstGeom);
										

									pic.appendChild(picnvPicPr);
									pic.appendChild(picblipFill);
									pic.appendChild(picspPr);
																			
									
								agraphicdata.appendChild(pic);

							agraphic.appendChild(agraphicdata);

							wpinline.appendChild(wpextent);
							wpinline.appendChild(wpeffectExtent);
							wpinline.appendChild(wpdocPr);
							wpinline.appendChild(wpcNvGraphicFramePr);							
							wpinline.appendChild(agraphic);

					drawEle.appendChild(wpinline);

					newEleImage.appendChild(drawEle);
								  

		return newEleImage;
		}		
		/**
		* function that check if string starts with prefix
		*/
		this.stringStartsWith=function (string, prefix) {
		    return string.slice(0, prefix.length) == prefix;
		}



		/** 
		*  función que forma el docx. Genera el zip y lo guarda
		*  @param: XmlDocumentDocx  el document.xml formado
		*  
		*/

		this.createDocx=function (XmlDocumentDocx, zip){

			
		     var relationShips = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
		  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'+
		'</Relationships>';

		
		     var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
		  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
		  '<Default ContentType="image/jpeg" Extension="jpg"/>'+
		'<Default ContentType="image/png" Extension="png"/>'+
		'<Default ContentType="image/gif" Extension="gif"/>'+
		  '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'+
		  '<Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" PartName="/word/styles.xml"/>'+
		  '<Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" PartName="/word/numbering.xml"/>'+
		'</Types>';

		
		var head_docx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">';

		var footer = '</w:document>';		 

		 // CARGAR ESTYLOS AQUI

		   var estilos = `
		  <w:styles xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15 wp14">
		<w:style w:type="paragraph" w:styleId="Heading1">
			<w:name w:val="Heading 1"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="48"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading2">
			<w:name w:val="Heading 2"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="38"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading3">
			<w:name w:val="Heading 3"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="35"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading4">
			<w:name w:val="Heading 4"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="30"/>
			</w:rPr>
			</w:style>	
		<w:style w:type="paragraph" w:styleId="Heading5">
			<w:name w:val="Heading 5"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="28"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading6">
			<w:name w:val="Heading 6"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="25"/>
			</w:rPr>
			</w:style>	
			<w:style w:type="paragraph" w:styleId="blockQuote">
			<w:name w:val="blockQuote"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>			
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>			
				<w:spacing w:before="360" w:after="360"/>
				<w:ind w:left="1080" w:right="1080" />
			</w:pPr>
			<w:rPr>				
				<w:color w:val="#222222"/>
				<w:sz w:val="30"/>
			</w:rPr>			
			</w:style>			
		<w:style xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="Hyperlink" mc:Ignorable="w14">
			<w:name xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="Hyperlink"/>
			<w:basedOn xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="DefaultParagraphFont"/>
			<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="99"/>
			<w:unhideWhenUsed xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
			<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:color w:val="0563C1" />
			<w:b/>
			<w:u w:val="single"/>
			</w:rPr>
			</w:style>
		<w:style mc:Ignorable="w14" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
  w:styleId="ListParagraph" w:type="paragraph">
			<w:name xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="List Paragraph"/>
			<w:basedOn xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="Normal"/>
			<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="34"/>
			<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
			<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:left="720"/>
			<w:contextualSpacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
			</w:pPr>
			</w:style>

		</w:styles>
		`;


		var cabeceraNumbering = '<?xml version="1.0" encoding="utf-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
		var footerNumbering = '</w:numbering>';
		
		numberingString = cabeceraNumbering+numberingString+numIdString+footerNumbering;
			
			// create a relationShips file
			zip.file("_rels/.rels", relationShips);
			// create a Content_tuypes.xml file
			zip.file("[Content_Types].xml", contentTypes);
			// document.xml	
			XmlDocumentDocx= head_docx+XmlDocumentDocx+footer;

			zip.file('word/document.xml', XmlDocumentDocx);
			zip.file('word/styles.xml',estilos);   
			zip.file('word/numbering.xml',numberingString);
				
			//var content = zip.generate({type:"blob"});

			// new lib jszip version
			zip.generateAsync({type:"blob"}).then(function(content) {
			    // see FileSaver.js
			    saveAs(content, "example.docx");
			});
			

		//return content;
		
		}		

	//  FUNCION parseHTMLtoDocx node
	this.parseHTMLtoDocx = function (node,xmlDoc,padreXML){  

	   	var  encabezados =['H1','H2','H3','H4','H5','H6'];
	  		
		// CREAR NODO BODY
	    if(node.nodeName === 'BODY'){

		    var newEle = xmlDoc.createElement("w:body");  
			
		}
		
	    // CREAR NODO ENCABEZADO
	    else if(encabezados.indexOf(node.nodeName) >= 0){ 
	    	
			var pos= encabezados.indexOf(node.nodeName) +1;	
			var newEle = this.createHeading(node,xmlDoc,pos);			

		}
		
		// CREAR NODO P.  
	    else if(node.nodeName === 'P'){

	    	var newEle = this.createNodeParagraphOrRun(node,xmlDoc);
			
		}

		// CREAR NODO STRONG
	    else if(node.nodeName === 'STRONG'){

			var newEle = this.createNodeStrong(node,xmlDoc);

		}

		// CREAR NODO EM
	    else if(node.nodeName === 'EM'){
	    				
			var newEle = this.createNodeEM(node,xmlDoc);

		}
		else if(node.nodeName === 'BR'){ 
		var newEle = this.createNodeBR(node,xmlDoc);

		}
		// CREAR NODO HIPERLINK
		else if( node.nodeName === 'A'){  
										
			    if (node.href ){  // EXTERNAL LINK
			    	
			    	var idLink = 'link'+numLink;
			    	relsDocumentXML = relsDocumentXML+'<Relationship Id="'+idLink+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'+node.href+'" TargetMode="External"/>';

			    	var newEle = this.createHiperlinkNode(node,xmlDoc,numLink);

				numLink++;	
			    }	
			    else var newEle = this.createNodeParagraphOrRun(node,xmlDoc);		    

		}
		
		// CREAR NODO LISTA . estos tags directamente se insertan en el body independientemente, en parrafos diferenciados
		else if( node.nodeName === 'UL' || node.nodeName === 'OL' ){ 

		// SI LISTA ANIDADA	. SIGO EN LA MISMA LISTA ITEM SE INCREMENTA EN UNO PARA EL SIGUIENTE NIVEL DE LA LISTA
		

		if(node.parentNode.nodeName != 'LI') countList++;

		if( node.nodeName === 'UL' )		numberingString = createAstractNumListBullet(numberingString,countList);
		else if(node.nodeName === 'OL' )    numberingString = createAstractNumListDecimal(numberingString,countList);
		
		numIdString = createNumIdList(numIdString, countList);

		var newEle = xmlDoc.createElement("w:Noinsert"); 

		}

		else if( node.nodeName === 'LI' ){  
			
			var levelList =0;
			var granFather = node.parentNode.parentNode;

			// CUANTOS ABUELOS TENGO
			if( granFather ){
				
				var abu = true;

				while(abu && granFather.nodeName === 'LI'){ levelList++; 
				
						if ( granFather.parentNode.parentNode ){  abu = true;  granFather = granFather.parentNode.parentNode; }
						else abu = false;					
				}
			}  
				
			var newEle = this.createNodeLi(node,xmlDoc,countList,levelList);
								
	    }

		// CREAR NODO TEXT. 
	    else if(node.nodeName === '#text'){ 

	    	var newEle = this.createMyTextNode(node,xmlDoc);
		    	
	    }
	    // CREAR NODO IMAGEN
	    else if( node.nodeName === 'IMG'){	    	

	    var format = '.png';
	    var nameFile = 'image'+numImg+format;  				
		var relashionImg = 'rId'+numImg;
		var imgEle;

		
		var dataImg = node.attributes[0].value;  
		

				if (!this.stringStartsWith(dataImg, 'data:')){  
					imgEle =  this.nodeVoid(xmlDoc);   
				}
	    		
	    		else { 

	    			var srcImg = dataImg.replace(/^data:image\/.+;base64,/,' '); 
	    			//GUARDO LA IMAGEN EN MEDIA
	    			zip.file("word/media/"+nameFile, srcImg , {base64: true});

	    			// CREO LA RELATIONSHIP DE LA IMAGEN CREADA
	    			relsDocumentXML = relsDocumentXML+'<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/'+nameFile+'" Id="'+relashionImg+'" />';			
		
	    			imgEle = this.createDrawingNodeIMG(node,dataImg,xmlDoc,numImg);	    			    			 

	    			numImg++;
	    		}


	    		if( node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD'){

	    			var newEle = xmlDoc.createElement("w:p");
	    			newEle.appendChild(imgEle);

	    		}
	    		else 	var newEle = imgEle;	    		

	    }
	    // IF TABLE
	    else if(  node.nodeName  === 'TABLE'){ 

	    var newEle = this.createTableNode(node,xmlDoc);

	    }
	    else if( node.nodeName === 'THEAD' ){ // no se inserta en el xml

	    	 var newEle =  xmlDoc.createElement('w:Noinsert');

	    }
	    else if( node.nodeName === 'TBODY' ){// no se inserta en el xml

	    	 var newEle =  xmlDoc.createElement('w:Noinsert');

	    }
	    else if( node.nodeName === 'TR' ){

		    var newEle = this.createNodeTR(node,xmlDoc);
		    // calculo num columnas
		    cols = 0;
		    for (var c=0;c < node.childNodes.length;c++){
		    	if( node.childNodes[c].nodeName === 'TD') cols++;
		    }
		    // SI FILA VACÍA.CONTROLAR!
		    if(cols ==  0){

		    	var eleTD = document.createElement("TD");  node.appendChild(eleTD); 

			}
		    

	    }
	    else if(  node.nodeName === 'TD' ){

	    var newEle = this.createNodeTD(node,xmlDoc,cols);
	    }
	    // BLOQUOTE
	    else if( node.nodeName === 'BLOCKQUOTE' ){
	    	
	    	var newEle = this.createNodeBlockquote(node,xmlDoc);
	    }
	  
	    // SI SE ENCUENTRA UN TAG QUE NO RECONOCE. 
	    else {   
	    	
	    	var newEle = this.createNodeParagraphOrRun(node,xmlDoc);
	    	
	    }
	    
	    
	     

	return newEle;
	}

	createNodeXML(nodeParent[0],xmlDoc,padreXML); // primer nodo es body[0]

	
	relsDocumentXML = relsDocumentXML+'<Relationship Id="Rnumbering1" Target="/word/numbering.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"/>';
	// cerrar rels Document.xml.rels
	relsDocumentXML = relsDocumentXML+'</Relationships>';   

	zip.file("word/_rels/document.xml.rels", relsDocumentXML);


	var xml = xmlDoc.getElementsByTagName('w:body');
	
	var content = this.createDocx(xml[0].outerHTML,zip);

return content;	
}










