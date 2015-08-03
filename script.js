var bload, input;
var argsList = [], resultList = [];
var argsHeader = [], blocksHeader = [];
var workbook, sheet;
var tableCont, table;


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
		for (var i in this.coords) {
			var c1 = this.coords[i], c2 = block2.coords[i];
			if (c2[0] >= c1[0] && c2[0] <= c1[1])
				result.coords.push([c2[0], Math.min(c1[1], c2[1])]);
			else if (c2[1] >= c1[0] && c2[1] <= c1[1])
				result.coords.push([Math.max(c1[0], c2[0]), c2[1]]);
			else
				return null;
		}
		return result;
	}
 } 

var fullOrderMin, fullOrderMax
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
	if (iter.toString() == fullOrderMin.toString())
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
	this.plus = [new Block(block)];
	this.minus = [];
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
			for (var j = frag[i].length - 1; j > 0; j--)
				if (frag[i][j] == frag[i][j - 1])
					frag[i].splice(j, 1);
		}

		var iterArr = [];
		fullOrderMax = [];
		fullOrderMin = [];
		for (var i in frag) {
			fullOrderMin.push(0);
			fullOrderMax.push(frag[i].length - 2);
		}
		var iter = fullOrderMin.slice(0);
		while (iter) {
			iterArr.push(iter.slice(0));
			iter = fullOrderInc(iter);
		}

		for (var i in this.minus) {
			fullOrderMax = [];
			fullOrderMin = [];
			for (var j in this.minus[i].coords) {
				fullOrderMin[j] = binIndexOf(frag[j], this.minus[i].coords[j][0]);
				fullOrderMax[j] = binIndexOf(frag[j], this.minus[i].coords[j][1]) - 1;
			}
			var iter = fullOrderMin.slice(0);
			while (iter) {
				var iterNumber = iter[iter.length - 1];
				for (var k = iter.length - 1; k > 0; k--)
					iterNumber = iterNumber * (frag[k].length - 1) + iter[k - 1];
				delete iterArr[iterNumber];
				iter = fullOrderInc(iter);
			}	
		}

		this.plus = [];
		this.minus = [];
		for (var i in iterArr) {
			var iter = iterArr[i];
			if (!iter)
				continue;
			var tmpArr = [];
			for (var j in iter)
				tmpArr.push([frag[j][iter[j]], frag[j][iter[j] + 1]]);
			this.plus.push(new Block(tmpArr));
		}
	}
	this.toString = function(){
		var res = ""
		for (var i in this.plus){
			if (i > 0)
				res += " and ";
			res += JSON.stringify(this.plus[i].coords);
		}

		var excl = ""
		for (var i in this.minus){
			if (i > 0)
				excl += " and ";
			excl += JSON.stringify(this.minus[i].coords);
		}

		if (excl.length > 0)
			res += " exclude (" + excl + ")";

		return res
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
	var range = XLSX.utils.decode_range(sheet['!ref']);
	if (!range)
		return 'no !ref in sheet';

	var baseArgs, endArgs, baseBlocks, endBlocks;
	loop: 
		for (var y = 0; y < 5; y++)
			for (var x = 0; x <= range.e.c; x++) {
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
		return 'Якоря не найдены. Расставьте #args, #args_end, #blocks, #blocks_end. Все якоря должны быть в одной строчке, между которой и первой строчкой данных расположены 2 строки заголовков.';
	if (baseArgs.r != baseBlocks.r)
		return 'Неверные позиции якорей.';

	argsHeader = [], blocksHeader = [];
	for (var y = baseArgs.r + 1; y <= baseArgs.r + 2; y++){
		var q = [];
		for (var x = baseArgs.c; x <= endArgs.c; x++){
			var v = getCell(sheet, x, y)
			q.push(v);
		}
		argsHeader.push(q);

		q = [];
		for (var x = baseBlocks.c; x <= endBlocks.c; x++){
			var v = getCell(sheet, x, y)
			q.push(v);
		}
		blocksHeader.push(q);
	}

	argsList = [], resultList = [];
	var yMin = baseArgs.r + 3; 
	loop1:
	for (var y = yMin; y <= range.e.r; y++) {
		var arr = [];
		for (var x = baseArgs.c; x <= endArgs.c; x++){
			var v = getCell(sheet, x, y);
			if(!v)
				break loop1;
			arr.push(v);
		}
		argsList.push(new Args(y - yMin, arr));
	}
	for (var y = yMin; y <= range.e.r; y++) {
		var arr = [];
		for (var x = baseBlocks.c; x < endBlocks.c; x += 2) {
			var min = getCell(sheet, x, y);
			var max = getCell(sheet, x + 1, y);
			if (typeof(min) != 'number' || typeof(max) != 'number' || min > max)
				return 'неверный формат данных в диапазоне ' + XLSX.utils.encode_cell({c : x, r : y}) + ':' + XLSX.utils.encode_cell({c : x + 1, r : y});
			arr.push([getCell(sheet, x, y), getCell(sheet, x + 1, y)]);
		}
		resultList.push(new BlocksSet(y - yMin, arr));
	}

	return true;
} 

function handleFile(e) {
	var files = input.files;
	if (!files || files.length < 1) {
		return "Вам стоит выбрать файл"
	}
	var f = files[0];
	var reader = new FileReader();
	var name = f.name;
	reader.onload = function(e) {
		var data = e.target.result;
		var workbook = XLSX.read(data, {type: 'binary'});
		prepareXls(workbook);
	};
	reader.readAsBinaryString(f);

	return true
}

function calc(){
	for (var i in resultList)
		for (var j in resultList)
			if (i != j) {
				var inter = resultList[i].plus[0].intersect(resultList[j].plus[0]);
				if (inter)
					resultList[i].minus.push(inter);
			}
	for (var i in resultList)
		resultList[i].applyMinus();
	console.log("resultList", resultList);

	printResults();
}

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

		jQuery("<td>").text(resultList[y].toString()).appendTo(row);

		table.append(row);
	}

	tableCont.append(table);

	//console.log(tableCont, table)
} 

function restart() {
	jQuery('.steps').hide();
	jQuery('.step1').show();
	jQuery('.error').text("").hide();
	jQuery('.step_restart').hide();
	if (table && table.length > 0) {
		table.remove();
	}
}

window.onload = function() {
	input = document.getElementById("myInput")
	jQuery('#btn_load').click(function(){
		var err = handleFile();
		if(err !== true){
			jQuery('.error').text("ошибка: " + err).show();
			return;
		}
		jQuery('.error').hide();
		jQuery('.steps').hide();
		jQuery('.step2').show();
		jQuery('.step_restart').show();
	});
	jQuery('#btn_parse').click(function(){
		err = parseXls();
		if(err !== true){
			jQuery('.error').text("ошибка парсинга файла: " + err).show();
			return;
		}
		jQuery('.error').hide();
		printResults();

		jQuery('.steps').hide();
		jQuery('.step3').show();
	});
	jQuery('#btn_calc').click(function(){
		calc();
		jQuery('.steps').hide();
	});
	jQuery('#btn_restart').click(function(){
		restart();
	});

	tableCont = jQuery("#table_cont");
	restart();
	console.log("init!")
}
