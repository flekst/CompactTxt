#target "Indesign"
/** CompactTXT - Скрипт для заполнения выбранных фреймов текстом.
    Пригодится в случае необходимости занять всё выделенное пространство, 
    когда кегль не имеет принципиального значения.
    Автоматически растягивает/уменьшает размер шрифта в фрейме  заполняя доступное пространство. 
    Работает в том числе со связанными фреймами. 
    Рекомендую назначить на горячую клавишу ALT+C - 
    она прекрасно коррелирует с комбинацией CTRL+ALT+C
 */
/**  Copyright (c) 2008-2014, Eugeny Borisov. All rights reserved. */

const CompactTXTerrorMessages = {
    StoryOverflowControl_MinTextPoinSize:    "Ошибка. Текст получается\nменьше минимально допустимого размера.",
    StoryOverflowControl_MaxTextPointSize:  "Ошибка. Текст получается\nмаксимально допустимого размера.",
    StoryOverflowControl_InfinityLoop:           "Ошибка. Не получается разместить текст в фрейм - идёт бесконечный цикл"
}

const c_myApp = app;
const c_myDoc = c_myApp.activeDocument;

const c_MinTextPointSize = 0;
const c_MaxTextPointSize = 65536;
const c_InitPointSize = 9;

const c_Max_loopsCount = 999;
var c_loopsCount = 0;

var c_currSelect = c_myDoc.selection;
var c_selCount = c_myDoc.selection.length;
var c_objects2Work=Array();
/** Добавляет элемент в массив, если он уникален */
Array.prototype.addunique = function(element) {
    var len = this.length;
    while (len--) { 
            if (this[len] == element) return; 
     }
     this.push (element);   
     
}
/** Вызывает callback-функцию для каждого элемента массива */
Array.prototype.map = function (callback) {
    var len = this.length;
    while (len--) { callback (this[len]);  }
}
/** Добавляет объект в массив, если он уникален */
Object.prototype.addToArrayIfUnique = function(targetArray) {  targetArray.addunique(this); }

/** c_objects2Work -- массив объектов с которыми нужно работать */
function CalcObjectForHandle(inputObject) {
	switch (String(inputObject)) {
			case "[object TextFrame]":
			case "[object Text]":
			case '[object InsertionPoint]':
			case '[object Character]': 
			case '[object Word]': 
			case '[object Line]': 
			case '[object Paragraph]': 
			case '[object TextColumn]':
                        inputObject = inputObject.parentStory;
  		   case "[object Story]" :
					inputObject.addToArrayIfUnique(c_objects2Work);
				break;                  
			case '[object Group]':
                        inputObject.allPageItems.everyItem().addToArrayIfUnique(c_objects2Work);
					break;
			default: break;
	}
}

/* Меняет значение на указанный процент с точностью в пять сотых */
function calcNewSize(oldSize, percent) {
        var newSize = oldSize*percent;
        if (newSize > 5)  newSize = Math.round(newSize*20)/20;
        return newSize;
}

/** Функция, изменяющая размер текста, на указанную единицу (в процентах), пока не поменяется статус overflow 
 если процент отрицательный - то ждем overflow==false, если положительный - true, 
//  Использует значения MinTextPoinSize и c_MaxTextPointSize.
  При выходе за их рамки функция генерирует исключения 'StoryOverflowControl_MinTextPoinSize' и 'StoryOverflowControl_MaxTextPointSize'
  При percent = 0 генерирует исключение 'StoryOverflowControl_zeroPercent';
*/
function StoryOverflowControl(inpStory, percent) {
    percent = (percent/100)+1;
    var newLeadingSize;
  
    if ( percent == 1) throw ('Ошибка в ДНК программиста');
    var waitForOverflow = (percent > 1)  ? true: false;

    if (inpStory.overflows == waitForOverflow) return true;

    var newPointSize = calcNewSize(inpStory.pointSize, percent);
  
    newLeadingSize  = ( inpStory.leading == Leading.auto ) ?  
                                newLeadingSize = Leading.auto :  calcNewSize(inpStory.leading, percent);
  
    var oldSize = newPointSize;
   
   /** основные "качели" - цил до изменения состояния Overflow */
    while (1) {
		inpStory.pointSize = newPointSize;
         if (newLeadingSize != Leading.auto ) {  
            inpStory.leading = newLeadingSize;
         }

		if (percent < 1 )  {
			inpStory.pointSize = newPointSize;
             inpStory.leading = newLeadingSize;
		}
	
		if (inpStory.overflows == waitForOverflow) return true;

        newPointSize = calcNewSize(inpStory.pointSize, percent);
        if (newLeadingSize != Leading.auto) {  
            newLeadingSize = calcNewSize(inpStory.leading, percent);
        }
        if (oldSize == newPointSize) break;

        oldSize = newPointSize;
        c_loopsCount++;
        if (c_loopsCount == c_Max_loopsCount) throw (CompactTXTerrorMessages.StoryOverflowControl_InfinityLoop);
        if (inpStory.pointSize < c_MinTextPointSize) throw (CompactTXTerrorMessages.StoryOverflowControl_MinTextPoinSize);
        if (inpStory.pointSize > c_MaxTextPointSize) throw (CompactTXTerrorMessages.StoryOverflowControl_MaxTextPointSize);
    }
    return false;
}

/** обработка одного текста */
function HandleStory (inputSory) {

    c_loopsCount = 0;            
    if (inputSory.leading != Leading.auto ) {  
        var newLeading = c_InitPointSize*inputSory.leading/inputSory.pointSize;
        inputSory.leading = newLeading;
    }
     inputSory.pointSize = c_InitPointSize;

    StoryOverflowControl(inputSory,+100);   StoryOverflowControl(inputSory,-50);
	StoryOverflowControl(inputSory,+25);    StoryOverflowControl(inputSory,-10);
	StoryOverflowControl(inputSory,+5);     StoryOverflowControl(inputSory,-2.5);
	StoryOverflowControl(inputSory,+1);     StoryOverflowControl(inputSory,-0.5);	
}

/** головная функция */
function main () {
    try {
       c_currSelect.map(CalcObjectForHandle);
       c_objects2Work.map(HandleStory);
    } catch (err) {
          alert (err); 
          exit(); 
      }
}

/** Возможность инклюда этого скрипта.
    Для того, что бы не выполнялась головная функция необходимо объявить переменную runCompactTXT и установить ей значение false */
try { runCompactTXT == undefined; } catch(err) { runCompactTXT = true; }

if (runCompactTXT) {
	 app.doScript(main, ScriptLanguage.JAVASCRIPT, [], UndoModes.FAST_ENTIRE_SCRIPT, 'compactTXT');
}
