var button, input;
var argsList = [], resultList = [];
var sheetG;


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

function BlocksSet(id, block) {
	this.id = id;
	this.plus = [new Block(block)]
	this.minus = []
	this.applyMinus = function() {
		var frag = [];
		for (var i = 0; i < this.plus[0].length; i++)
			frag.push({});
		for (var b in this.minus) {
			for (var c in this.minus[b].coords) {
				var d = this.minus[b].coords[c];
			}
		}
	}
}

function getCell(sheet, x, y) {
	var cell = sheet[XLSX.utils.encode_cell({c : x, r : y})];
	if (!cell)
		return null;
	return cell.v;
}

function parseXls(workbook) {
	var sheet = workbook.Sheets[workbook.SheetNames[2]];
	var range = XLSX.utils.decode_range(sheet['!ref'])
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

	var yMin = baseArgs.r + 3; 
	for (var y = yMin; y < range.e.r; y++) {
		var arr = []
		for (var x = baseArgs.c; x <= endArgs.c; x++)
			arr.push(getCell(sheet, x, y));
		argsList.push(new Args(y - yMin, arr));
	}
	for (var y = yMin; y < range.e.r; y++) {
		var arr = []
		for (var x = baseBlocks.c; x < endBlocks.c; x += 2)
			arr.push([getCell(sheet, x, y), getCell(sheet, x + 1, y)]);
		resultList.push(new BlocksSet(y - yMin, arr));
	}
	
	sheetG = sheet
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
		parseXls(workbook);
	};
	reader.readAsBinaryString(f);
}



window.onload = function() {
	input= document.getElementById("myInput")
	button = document.getElementById("myButton")
	button.addEventListener('click', handleFile, false)
	console.log("init!")
}
