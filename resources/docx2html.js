function convertDocX(attr){

	if(!attr) attr = {};
	if(attr.onload && typeof attr.onload==="function") this.onload = attr.onload;
	if(attr.onselect && typeof attr.onselect==="function") this.onselect = attr.onselect;

	this.relations = {};
	this.input = "";
	this.header = "";
	this.images = new Array();


	return this;
}

convertDocX.prototype.relationships = function(data){
	console.log('relationships',data)
	var _obj = this;
	
	data.replace(/<Relationship Id="([^\"]+)" Type="([^\"]+)" Target="([^\"]+)" TargetMode="([^\"]+)"\/>/g,function(m,p1,p2,p3,p4){
		_obj.relations[p1] = p3;
		return "";
	});
	
	return this;
}

convertDocX.prototype.load = function(data){

	this.input = data;
	var header = this.header;

	var _obj = this;
	
	function cleanMS(s){
		// smart single quotes and apostrophe
		s = s.replace(/[\u2018\u2019\u201A]/g, "\'");

		// smart double quotes
		s = s.replace(/[\u201C\u201D\u201E]/g, "\"");

		// ellipsis
		s = s.replace(/\u2026/g, "...");

		// dashes
		s = s.replace(/[\u2013\u2014]/g, "-");
		return s;
	}
	function removeEmptyTags(str){
		return str.replace(/<([^\>\s]+) ?[^\>]*>[\n\r\t\s]*<\/\1>/g,"");
	}
	function makeHyperlinks(str){
		/*<w:hyperlink r:id=\"rId11\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" w:history=\"1\"><w:r w:rsidR=\"00BD4E61\" w:rsidRPr=\"00BD4E61\"><w:rPr><w:rStyle w:val=\"Hyperlink\"/></w:rPr><w:t>https://docs.google.com/forms/d/e/1FAIpQLSdDrytqID_-0qhB47jq-NO2wWcHT427KPiRufA1mygv7ZMP2g/viewform?c=0&amp;w</w:t></w:r></w:hyperlink>*/
		return str.replace(/<w:hyperlink r:id="([^\"]+)".*?<w:t>([^\<]*)<\/w:t><\/w:r><\/w:hyperlink>/g,function(m,p1,p2){
			return "["+p2+"]("+(_obj.relations[p1] ? _obj.relations[p1] : p1)+")";
		});
	}
	function makeImages(str){
		var i = -1;
		// <w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="7BB2C9B8" wp14:editId="14F17FC6"><wp:extent cx="1119257" cy="1119257"/><wp:effectExtent l="0" t="0" r="5080" b="0"/><wp:docPr id="7" name="Picture 7" descr="odileeds-autumn-250.png"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="Picture 3" descr="odileeds-autumn-250.png"/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId6"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="1136330" cy="1136330"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>
		// <w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="2094B4CA" wp14:editId="03E86F98"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>4914900</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-220980</wp:posOffset></wp:positionV><wp:extent cx="828675" cy="797560"/><wp:effectExtent l="0" t="0" r="9525" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="27" name="Picture 27"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="universitylogo.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="828675" cy="797560"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic><wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom="page"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV></wp:anchor></w:drawing>
		return str.replace(/<w:drawing>.*?<wp:docPr id="([^\"]+)" name="([^\"]+)"[^\>]+>.*?<\/w:drawing>/g,function(m,p1,p2){
			console.log(m,p1,p2,_obj,i)
			i++;
			return '![image]('+_obj.images[i]+' "'+p1+'")';
		});
	}

	var mode = function mode(arr) {
		var numMapping = {};
		var greatestFreq = 0;
		var mode;
		arr.forEach(function findMode(number) {
			numMapping[number] = (numMapping[number] || 0) + 1;

			if (greatestFreq < numMapping[number]) {
				greatestFreq = numMapping[number];
				mode = number;
			}
		});
		return +mode;
	}
	

	header = header.replace(/[\n\r]/g,"");
	data = data.replace(/[\n\r]/g,"");

	header = header.replace(/^.*<w:hdr[^\>]+>(.*)<\/w:hdr>/,function(m,p1){ return p1; });
console.log('header',header)
	data = data.replace(/<w:body>/,function(m,p1){ return "<w:body>"+header; });
console.log('new',data)

	// Remove dodgy Microsoft Word characters
	data = cleanMS(data);
	console.log('relations',this.relations);
	data = makeImages(data);
	data = makeHyperlinks(data);
	data = data.replace(/<w:tab\/>/g," ");


	// Find all used font sizes
	tmp = data;
	bits = tmp.split(/<w:sz /);
	sizes = {};
	size = new Array();
	for(var s = 1; s < bits.length; s++){
		m = (bits[s].match(/w:val=\"([0-9]+)\"/));
		if(m.length==2){
			if(!sizes[m[1]]) sizes[m[1]] = 0;
			sizes[m[1]]++;
			size.push(m[1]);
		}
	}
	normal = mode(size);
//			console.log(sizes,size,mode(size));

	// Convert to DOM
	parser = new DOMParser();
	xmlDoc = parser.parseFromString(data, "text/xml");
	body = xmlDoc.childNodes[0].childNodes[0];
	html = "";
	// Loop over each item
	for(var i = 0; i < body.childNodes.length; i++){
		var type = "p";
		var out = "";

		content = body.childNodes[i].innerHTML.replace(/<[^\>]+>/g,"");
		trimmed = content.replace(/[\n\r\t\s]/g,"");

		if(body.childNodes[i].nodeName == "w:tbl") type = "table";
		if(body.childNodes[i].innerHTML.indexOf("<w:pStyle w:val=\"ListParagraph\"/>") >= 0) type = "li";
		if(body.childNodes[i].innerHTML.indexOf("<w:numPr>") >= 0) type = "li";

		if(type == "p"){
			m = (body.childNodes[i].innerHTML.match(/sz w:val=\"([0-9]+)\"/));
			fontsize = normal;
			if(m && m.length==2){
				fontsize = parseInt(m[1]);
				if(fontsize > normal) type = "h2";
			}
			if(type == "p" && fontsize >= normal){
				if(body.childNodes[i].childNodes.length > 0){
					h = body.childNodes[i].childNodes[0].innerHTML;
					if(h){
						if(h.indexOf("<w:b") >= 0) type = "h3";
					}
					//console.log(i,type,content,fontsize)
				}
			}
			if(type == "p" && body.childNodes[i].innerHTML.indexOf(/\u2022/) >= 0) type = "li";
			//console.log(i,type,content,body.childNodes[i])
		}


		if(trimmed){

			if(type == "table"){
				table = body.childNodes[i];
				out += "<table>";
				row = 0;
				for(var t = 0; t < table.childNodes.length; t++){
					tr = table.childNodes[t];
				
					if((tr.nodeName == "w:tr")){
						row++;
						if(row == 1) out += "<thead>";
						if(row == 2) out += "<tbody>";
						out += "<tr>";
						//console.log(tr);
						for(var c = 0; c < tr.childNodes.length; c++){
							out += (row > 1 ? "\t<td>":"\t<th>")+(tr.childNodes[c].innerHTML.replace(/<[^a][^\>]+>/g,""))+(row > 1 ? "</td>":"</th>");
						}
						out += "</tr>";
						if(row == 1) out += "</thead>";
					}
				}
				out += "</tbody></table>";
		
			}else{
				// remove trailing spaces
				content = content.replace(/\s*$/,"");
				// Build output string
				out = "<"+type+">"+content+"</"+type+">"; 
			}

			//console.log(i,type,body.childNodes[i],out,content);
			html += out;
		}
	}



	// Remove empty paragraphs
	/*<w:p w14:paraId="70D0A2AE" w14:textId="77777777" w:rsidR="00363F74" w:rsidRDefault="00363F74" w:rsidP="00FB7A0A">
		<w:pPr>
			<w:rPr>
				<w:b />
			</w:rPr>
		</w:pPr>
	</w:p>*/

	// List
	/* <w:pPr>
		<w:pStyle w:val="ListParagraph" />
		<w:numPr>
			<w:ilvl w:val="1" />
			<w:numId w:val="15" />
		</w:numPr>
		<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" />
		<w:spacing w:after="0" w:line="240" w:lineRule="auto" />
		<w:rPr>
			<w:rFonts w:ascii="Arial" w:eastAsia="Times New Roman" w:hAnsi="Arial" w:cs="Arial" />
			<w:sz w:val="19" />
			<w:szCs w:val="19" />
			<w:lang w:eastAsia="en-GB" />
		</w:rPr>
	</w:pPr>
	
	
	<w:pPr>
		<w:pStyle w:val="ListParagraph" />
		<w:numPr>
			<w:ilvl w:val="0" />
			<w:numId w:val="16" />
		</w:numPr>
		<w:autoSpaceDE w:val="0" />
		<w:autoSpaceDN w:val="0" />
		<w:adjustRightInd w:val="0" />
		<w:spacing w:after="10" w:line="240" w:lineRule="auto" />
		<w:rPr>
			<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" />
			<w:color w:val="000000" />
		</w:rPr>
	</w:pPr>
	*/


	// Images are within paragraphs
	/* 
	<w:drawing>
			<wp:inline distT="0" distB="0" distL="0" distR="0" wp14:anchorId="126597CB" wp14:editId="5E7D0A54">
				<wp:extent cx="3522135" cy="1981200" />
				<wp:effectExtent l="0" t="0" r="2540" b="0" />
				<wp:docPr id="3" name="Picture 3" descr="https://pbs.twimg.com/media/C8EZ6drWkAAPy4u.jpg" />
				<wp:cNvGraphicFramePr>
					<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
				</wp:cNvGraphicFramePr>
				<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
					<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
						<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
							<pic:nvPicPr>
								<pic:cNvPr id="0" name="Picture 1" descr="https://pbs.twimg.com/media/C8EZ6drWkAAPy4u.jpg" />
								<pic:cNvPicPr>
									<a:picLocks noChangeAspect="1" noChangeArrowheads="1" />
								</pic:cNvPicPr>
							</pic:nvPicPr>
							<pic:blipFill>
								<a:blip r:embed="rId10" cstate="print">
									<a:extLst>
										<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
											<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0" />
										</a:ext>
									</a:extLst>
								</a:blip>
								<a:srcRect />
								<a:stretch>
									<a:fillRect />
								</a:stretch>
							</pic:blipFill>
							<pic:spPr bwMode="auto">
								<a:xfrm>
									<a:off x="0" y="0" />
									<a:ext cx="3527084" cy="1983984" />
								</a:xfrm>
								<a:prstGeom prst="rect">
									<a:avLst />
								</a:prstGeom>
								<a:noFill />
								<a:ln>
									<a:noFill />
								</a:ln>
							</pic:spPr>
						</pic:pic>
					</a:graphicData>
				</a:graphic>
			</wp:inline>
		</w:drawing>
	*/




	// New page
	/*		<w:p w14:paraId="03AB5883" w14:textId="5F6F3C35" w:rsidR="00363F74" w:rsidRDefault="00363F74">
		<w:pPr>
			<w:rPr>
				<w:b />
			</w:rPr>
		</w:pPr>
		<w:r>
			<w:rPr>
				<w:b />
			</w:rPr>
			<w:br w:type="page" />
		</w:r>
	</w:p>*/
	
	
	
	
	
	// Any <w:rPr> that contains <w:b /> or <w:sz info is probably a heading
	/*<w:rPr>
			<w:b />
			<w:sz w:val="28" />
			<w:szCs w:val="28" />
		</w:rPr>*/
	// If <w:pRr> contains bold or sizing, it is probably a heading
	/* <w:pPr>
		<w:jc w:val="center" />
		<w:rPr>
			<w:b />
			<w:sz w:val="28" />
			<w:szCs w:val="28" />
		</w:rPr>
	</w:pPr>*/

	// However this style defines a default and isn't a heading
	/*<w:pPr>
		<w:pStyle w:val="Default" />
		<w:rPr>
			<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cs="Century Gothic" />
			<w:sz w:val="22" />
			<w:szCs w:val="22" />
		</w:rPr>
	</w:pPr>
	
	Basically, use <w:pStyle w:val="Default" /> to find the default font size. Any
	change in font size is probably a hint of a heading. Or if the text goes bold.
	*/
	





/*
	// Remove the head and foot
	data = data.replace(/^.*<w:body>/,"").replace(/<\/w:body><\/w:document>.*$/,"");

	data = removeEmptyTags(data);
	data = data.replace(/<w:br ?\/>/g,"\n");
	data = data.replace(/<w:b ?\/>/g,"<BOLD>");
	data = data.replace(/<w:tab\/>/g,"\t");
	data = data.replace(/<w:r><w:rPr>\n<\/w:rPr>\t<\/w:r>/g,"\n\nREMOVED A\n\n");
	data = data.replace(/<w:r [^\>]*><w:rPr>\n<\/w:rPr>\t<\/w:r>/g,"\n\nREMOVED B\n\n");
	data = data.replace(/<w:rPr>\n<\/w:rPr>/g,"\n\nREMOVED C\n\n");
	data = data.replace(/<w:pPr>\n\nREMOVED[^\n]*\n\n<\/w:pPr>/g,"\n\nREMOVED D\n\n");
	data = removeEmptyTags(data);
	
	
	data = data.replace(/\n/g,"==NEWLINE==");
	data = data.replace(/<w:p.*>(.*)?<\w:p>/g,function(match,p1){ return "<p>"+p1+"</p>"; });
*/
	// REMOVE ALL TAGS
	//data = data.replace(/<[^\>]+>/g,"");
	data = data.replace(/\n\nREMOVED[^\n]*\n\n/g,"\n");

	// Convert MarkDown images back to HTML
	//![image](base64 "Title")
	html = html.replace(/!\[[^\]]+\]\(([^\s]+) \"([^\"]+)\"\)/g,function(m,p1,p2){ return '<img src="data:image/png;base64,'+p1+'" title="'+p2+'" />'; })
	html = html.replace(/<h2>((<img [^\>]+\/>)+)<\/h2>/g,function(m,p1){ return "<figure>\n\t"+p1+"\n</figure>"; });

	// Convert MarkDown links back to HTML
	html = html.replace(/\[([^\]]+)\]\(([^\)]+)\)/g,function(m,p1,p2){ return '<a href="'+p2+'">'+p1+'</a>'; })

	// Tidy
	html = html.replace(/(p|h[0-9]|table|figure)><li>/g,function(m,p){ return p+"><ul>\n<li>"; }).replace(/<\/li>[\n\t\s]*<(p|h[0-9]|table|figure)/g,function(m,p){ return "</li></ul>\n<"+p; });
	html = html.replace(/<\/([^\>]*)>/g,function(m,p){ return "<\/"+p+">\n";});
	html = html.replace(/<li>/g,"\t<li>");
	html = html.replace(/<li>\u2022 ?/g,"<li>");

	html = html.replace(/<(table|thead|tbody|tr)>/g,function(m,p){ return "<"+p+">\n";});
	html = html.replace(/<h/g,"\n\n<h");
	html = html.replace(/<p/g,"\n<p");
	html = html.replace(/<figure/g,"\n<figure");
	html = html.replace(/\\\* ARABIC ([0-9])/g,function(m,p){ return '&#'+(48+parseInt(p))+';'; } );
		
	for(var i = 0; i < this.images.length; i++){
	//	html = 'TEST <img src="data:image/png;base64,'+this.images[i]+'" />'+html;
	}
	if(typeof this.onload==="function") this.onload.call(this,input,html);	

	return this;
}
convertDocX.prototype.handleFileSelect = function(evt,typ){

	evt.stopPropagation();
	evt.preventDefault();
	if(typeof this.onselect==="function") this.onselect.call(this);
	
	var files;
	if(evt.dataTransfer && evt.dataTransfer.files) files = evt.dataTransfer.files; // FileList object.
	if(!files && evt.target && evt.target.files) files = evt.target.files;

	function niceSize(b){
		if(b > 1e12) return (b/1e12).toFixed(2)+" TB";
		if(b > 1e9) return (b/1e9).toFixed(2)+" GB";
		if(b > 1e6) return (b/1e6).toFixed(2)+" MB";
		if(b > 1e3) return (b/1e3).toFixed(2)+" kB";
		return (b)+" bytes";
	}

	if(typ == "docx"){

		// files is a FileList of File objects. We will only look at the first
		f = files[0];

		var output = "";
		output += '<div><strong>'+ escape(f.name)+ '</strong> - ' + niceSize(f.size) + ', last modified: ' + (f.lastModified ? (new Date(f.lastModified)).toLocaleDateString() : 'n/a') + '</div>';

		var _obj = this;
		JSZip.loadAsync(f,{'compression':'DEFLATE'})                                   // 1) read the Blob
		.then(function(zip) {
			var dateAfter = new Date();
			var zipRels = zip.files["word/_rels/document.xml.rels"];
			var zipEntry = zip.files["word/document.xml"];
			console.log('jszip',zip.files["word/document.xml"],zip.files["word/_rels/document.xml.rels"]);

			// Lets extract any images
			// First build an array of image file names
			var icount = 0;
			var i = 0;
			var images = new Array();
			for(var f in zip.files){
				if(f.indexOf("word/media/")==0) images.push(f);
			}
			// Function to load an image
			function getImage(){
				f = images[icount];
				i = parseInt(f.replace(/[^0-9]/g,""))-1;
				console.log('getImage',i,f)
				zip.files[f].async("base64").then(function (content){
					console.log('base64',icount,i,content);
					_obj.images[i] = content;
					icount++;
					if(icount < images.length) getImage();
					else loadDoc();
				});
			}
			function loadRelationships(){
				// Load the relationships
				if(zip.files["word/_rels/document.xml.rels"]){
					console.log('loadRelationships')
					zip.files["word/_rels/document.xml.rels"].async("string").then(function (content) {
						_obj.relationships(content);
						loadHeader();
					});
				}else{
					loadHeader();
				}
			}
			function loadHeader(){
				if(zip.files["word/header1.xml"]){
					console.log('loadHeader');
					zip.files["word/header1.xml"].async("string").then(function (content) {
						//_obj.relationships(content);
						console.log(content)
						_obj.header += content;
						loadMain();
					});
				}else{
					loadMain();
				}
			}
			function loadMain(){
				console.log('loadMain')
				// Now process the rest
				zip.files["word/document.xml"].async("string").then(function (content) {
					console.log(content);
					_obj.load(content);
				});
			}

			// Function to load the rest of the document
			function loadDoc(){
				_obj.input = "";
				_obj.header = "";
				loadRelationships();
			}

			if(images.length > 0) getImage();	// This will call loadDoc when it has finished
			else loadDoc();
			

		}, function (e) {

		});

		S('#drop_zone').addClass('loaded').html(output);
		S('.step1').addClass('checked');
	}
	return this;
}
