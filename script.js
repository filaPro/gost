var bload, input;
var argsList = [], resultList = [];
var argsHeader = [], blocksHeader = [];
var workbook, sheet;
var tableCont, table


function Args(id, arr) {
	this.id = id;
	this.arr = arr;
}

function Block(arr) { 
	this.coords = arr;
	this.intersect = function(block2) {
		if (this.coords.length != block2.coords.length)
			return null;
		var result = new Block([]);
		var c1 = this.coords[i], c2 = block2.coords[i];
		for (var i = 0; i < this.coords.length; i++) {
			if (c2[0] >= c1[0] && c2[0] <= c1[1])
				result.coords.push([c2[0], Math.min(c1[0], c2[1])]);
			else if (c2[1] >= c1[0] && c2[1] <= c1[1])
				result.coords.push([Math.max(c1[0], c2[0]), c2[1]]);
			else
				return null;
		}
	}
} 

var fullOrderMin = [], fullOrderMax = []
function fullOrderInc(iter) {
	var i = 0;
	while (i < iter.length) {
		iter[i]++;
		if (iter[i] > fullOrderMax[i]) {
			iter[i] = fullOrderMin[i];
			i++;
			continue;
		}
		break;
	}
	if (iter == fullOrderMin)
		return null;
	return iter;
}

function binIndexOf(arr, searchElement) {
	var minIndex = 0;
	var maxIndex = arr.length - 1;
	var currentIndex;
	var currentElement;

	while (minIndex <= maxIndex) {
		currentIndex = (minIndex + maxIndex) / 2 | 0;
		currentElement = arr[currentIndex];

		if (currentElement < searchElement) {
			minIndex = currentIndex + 1;
		}
		else if (currentElement > searchElement) {
			maxIndex = currentIndex - 1;
		}
		else {
			return currentIndex;
		}
	}

	return -1;
}

 function BlocksSet(id, block) {
	this.id = id;
	this.plus = [new Block(block)]
	this.minus = []
 	this.applyMinus = function() {
		var frag = [];
		for (var i in this.plus[0].coords) {
			frag.push([]);
			frag[i].push(this.plus[0].coords[i][0]);
			frag[i].push(this.plus[0].coords[i][1]);
		}
		for (var b in this.minus) {
			for (var c in this.minus[b].coords) {
				var d = this.minus[b].coords[c];
				frag[c].push(d[0]);
				frag[c].push(d[1]);
			}
		}
		for (var i in frag) {
			frag[i].sort();
			for (var j = frag[i][frag[i].length - 1]; j > 0; j--)
				if (frag[i][j] == frag[i][j - 1])
					frag[i].splice(j, 1);
		}

		var iterArr = [];
		fullOrderMax = fullOrderMin = []
		for (var i in frag) {
			fullOrderMin[i] = 0;
			fullOrderMax[i] = frag[i].length - 2;
		}
		var iter = fullOrderMin;
		while (iter) {
			//var tmpArr = []
			//for (var i in iter.length)
			//	tmpArr.push([frag[i][iter[i]], frag[i][iter[i + 1]]]);
			//this.result.push(new Block(tmpArr));
			iterArr.push(iter);
			iter = fullOrderInc(iter);
		}

		for (var i in this.minus) {
			fullOrderMax = fullOrderMin = []
			for (var j in this.minus[j].coords) {
				fullOrderMin[j] = binIndexOf(frag[j], this.minus[i].coords[j][0]);
				fullOrderMax[j] = binIndexOf(frag[j], this.minus[i].coords[j][1] - 1);
			}
			var iter = fullOrderMin;
			while (iter) {
				var iterNumber = 0;
				for (var k in frag)
					iterNumber += iter[k] * Math.pow(frag[k].length, iter.length - k - 1);
				delete iterArr[iterNumber];
				iter = fullOrderInc(iter);
			}	
		}

		this.plus = []
		for (var i in iterArr) {
			var iter = iterArr[i];
			if (!iter)
				continue;
			var tmpArr = [];
			for (var j in iter.length)
				tmpArr.push([frag[j][iter[j]], frag[j][iter[j + 1]]]);
			this.plus.push(new Block(tmpArr));
		}
	}
}

function getCell(sheet, x, y) {
	var cell = sheet[XLSX.utils.encode_cell({c : x, r : y})];
	if (!cell)
		return null;
	return cell.v;
}

function prepareXls(w){
	workbook = w;
	var sel = jQuery("#sheet_id");
	sel.empty();
	for (var i in w.SheetNames)
		jQuery("<option>").val(i).text(w.SheetNames[i]).appendTo(sel);
}
function parseXls() {
	snum = jQuery("#sheet_id").val();
	var sheet = workbook.Sheets[workbook.SheetNames[snum]];
	console.log(sheet, snum);
	var range = XLSX.utils.decode_range(sheet['!ref']);
	if (!range)
		return 'no !ref in sheet';

	var baseArgs, endArgs, baseBlocks, endBlocks;
	loop: 
		for (var y = 0; y < 5; y++)
			for (var x = 0; x < range.e.c; x++) {
				var v = getCell(sheet, x, y);
				if (!v)
					continue;
				if (v == '#args')
					baseArgs = {c : x, r : y};
				if (v == '#blocks')
					baseBlocks = {c : x, r : y};
				if (v == '#args_end')
					endArgs = {c : x, r : y};
				if (v == '#blocks_end') {
					endBlocks = {c : x, r : y};
					break loop;
				}
			}
	if (!baseArgs || !endArgs || !baseBlocks || !endBlocks)
		return 'anchors not found';
	if (baseArgs.r != baseBlocks.r)
		return 'wrong positions of anchors';

	argsHeader = [], blocksHeader = [];
	for (var y = baseArgs.r+1; y <= baseArgs.r+2; y++){
		var q = [];
		for (var x = baseArgs.c; x <= endArgs.c; x++){
			var v = getCell(sheet, x, y)
			q.push(v);
			console.log(sheet, x, y, v)
		}
		argsHeader.push(q);

		q = [];
		for (var x = baseBlocks.c; x <= endBlocks.c; x++){
			var v = getCell(sheet, x, y)
			q.push(v);
		}
		blocksHeader.push(q);
	}

	var yMin = baseArgs.r + 3; 
	loop1:
	for (var y = yMin; y < range.e.r; y++) {
		var arr = [];
		for (var x = baseArgs.c; x <= endArgs.c; x++){
			var v = getCell(sheet, x, y);
			if(!v)
				break loop1;
			arr.push(v);
		}
		argsList.push(new Args(y - yMin, arr));
	}
	for (var y = yMin; y < range.e.r; y++) {
		var arr = [];
		for (var x = baseBlocks.c; x < endBlocks.c; x += 2)
			arr.push([getCell(sheet, x, y), getCell(sheet, x + 1, y)]);
		resultList.push(new BlocksSet(y - yMin, arr));
	}
	for (var i in resultList)
		resultList[i].applyMinus();

	sheetG = sheet;
	console.log(sheet, resultList, argsList);
	return true;
} 

function handleFile(e) {
	var files = input.files;
	console.log("open files", files);
	var f = files[0];
	var reader = new FileReader();
	var name = f.name;
	reader.onload = function(e) {
		var data = e.target.result;
		var workbook = XLSX.read(data, {type: 'binary'});
		prepareXls(workbook);
	};
	reader.readAsBinaryString(f);
}

function loadData(){
	err = parseXls();
	if(err !== true){
		alert("parsing error:" + err);
		return;
	}

	printResults();
}
function calc(){}

function printResults(){
	if (table && table.length > 0) {
		table.remove();
	}
	table = jQuery("<table>");
	for (var y in argsHeader){
		row = jQuery("<tr>");
		for (var x in argsHeader[y])
			jQuery("<th>").text(argsHeader[y][x]).appendTo(row);

		if (y == 0){
			var str = "";
			for (var x in blocksHeader[y]){
				if (!blocksHeader[y][x])
					continue;
				if(x > 0)
					str += " x ";
				str += "\"" + blocksHeader[y][x] + "\"";
			}
			jQuery("<th>").text(str).appendTo(row);
		}

		table.append(row);
	}

	for (var y in argsList) {
		var a = argsList[y];
		row = jQuery("<tr>");
		for (var x in a.arr)
			jQuery("<td>").text(a.arr[x]).appendTo(row);

		jQuery("<td>").text(JSON.stringify(resultList[y])).appendTo(row);

		table.append(row);
	}

	tableCont.append(table);

	console.log(tableCont, table)
}

window.onload = function() {
	input= document.getElementById("myInput")
	bload = document.getElementById("btn_load")
	bparse = document.getElementById("btn_parse")
	bcalc = document.getElementById("btn_calc")
	bload.addEventListener('click', handleFile, false)
	bparse.addEventListener('click', loadData, false)
	bcalc.addEventListener('click', calc, false)

	tableCont = jQuery("#table_cont")
	console.log("init!")
}
