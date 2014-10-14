#target "Indesign"
/** CompactTXT - Скрипт для заполнения выбранных фреймов текстом.
    Пригодится в случае необходимости занять всё выделенное пространство, 
    когда кегль не имеет принципиального значения.
    Автоматически растягивает/уменьшает размер шрифта в фрейме  заполняя доступное пространство. 
    Работает в том числе со связанными фреймами. 
    Рекомендую назначить на горячую клавишу ALT+C - 
    она прекрасно коррелирует с комбинацией CTRL+ALT+C
 */

/**  Copyright (c) 2008-2014, Eugeny Borisov a.k.a. kstati. All rights reserved.
Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation 
and/or other materials provided with the distribution.
Neither the name of the <ORGANIZATION> nor the names of its contributors may be used to endorse or promote products derived from this software
without specific prior written permission.
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND 
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES 
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; 
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND 
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT 
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, 
EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. */

/** 
    Перевод лицензии:
    Copyright (c) 2008-2014, Евгений Борисов известный как kstati.
Разрешается повторное распространение и использование как в виде исходного кода, так и в двоичной форме, с изменениями или без, 
при соблюдении следующих условий:
При повторном распространении исходного кода должно оставаться указанное выше уведомление об авторском праве, этот список условий 
и последующий отказ от гарантий.
При повторном распространении двоичного кода должна сохраняться указанная выше информация об авторском праве, этот список условий 
и последующий отказ от гарантий в документации и/или в других материалах, поставляемых при распространении.
Ни название <Организации>, ни имена ее сотрудников не могут быть использованы в качестве поддержки или продвижения продуктов, 
основанных на этом ПО без предварительного письменного разрешения.
ЭТА ПРОГРАММА ПРЕДОСТАВЛЕНА ВЛАДЕЛЬЦАМИ АВТОРСКИХ ПРАВ И/ИЛИ ДРУГИМИ СТОРОНАМИ «КАК ОНА ЕСТЬ» 
БЕЗ КАКОГО-ЛИБО ВИДА ГАРАНТИЙ, ВЫРАЖЕННЫХ ЯВНО ИЛИ ПОДРАЗУМЕВАЕМЫХ, ВКЛЮЧАЯ, НО НЕ ОГРАНИЧИВАЯСЬ ИМИ, 
ПОДРАЗУМЕВАЕМЫЕ ГАРАНТИИ КОММЕРЧЕСКОЙ ЦЕННОСТИ И ПРИГОДНОСТИ ДЛЯ КОНКРЕТНОЙ ЦЕЛИ. 
НИ В КОЕМ СЛУЧАЕ НИ ОДИН ВЛАДЕЛЕЦ АВТОРСКИХ ПРАВ И НИ ОДНО ДРУГОЕ ЛИЦО, 
КОТОРОЕ МОЖЕТ ИЗМЕНЯТЬ И/ИЛИ ПОВТОРНО РАСПРОСТРАНЯТЬ ПРОГРАММУ, КАК БЫЛО СКАЗАНО ВЫШЕ, 
НЕ НЕСЁТ ОТВЕТСТВЕННОСТИ, ВКЛЮЧАЯ ЛЮБЫЕ ОБЩИЕ, СЛУЧАЙНЫЕ, СПЕЦИАЛЬНЫЕ ИЛИ ПОСЛЕДОВАВШИЕ УБЫТКИ, 
ВСЛЕДСТВИЕ ИСПОЛЬЗОВАНИЯ ИЛИ НЕВОЗМОЖНОСТИ ИСПОЛЬЗОВАНИЯ ПРОГРАММЫ 
(ВКЛЮЧАЯ, НО НЕ ОГРАНИЧИВАЯСЬ ПОТЕРЕЙ ДАННЫХ, ИЛИ ДАННЫМИ, СТАВШИМИ НЕПРАВИЛЬНЫМИ, 
ИЛИ ПОТЕРЯМИ ПРИНЕСЕННЫМИ ИЗ-ЗА ВАС ИЛИ ТРЕТЬИХ ЛИЦ, ИЛИ ОТКАЗОМ ПРОГРАММЫ РАБОТАТЬ 
СОВМЕСТНО С ДРУГИМИ ПРОГРАММАМИ), ДАЖЕ ЕСЛИ ТАКОЙ ВЛАДЕЛЕЦ ИЛИ ДРУГОЕ ЛИЦО БЫЛИ ИЗВЕЩЕНЫ 
О ВОЗМОЖНОСТИ ТАКИХ УБЫТКОВ. */



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

/** Добавляю уникальные объекты в массив  */
function AddToArrayUnique(obj,Iarray) {
 var i = Iarray.length; 
 while (i--) if (Iarray[i] == obj) return;
 Iarray.push(obj);
}

/** c_objects2Work -- массив объектов с которыми нужно работать */
function CalcObjectForHandle(inputObject) {
	switch (String(inputObject)) {
			case "[object Story]" :
					AddToArrayUnique(inputObject, c_objects2Work);
				break;
			case "[object TextFrame]":
			case "[object Text]":
			case '[object InsertionPoint]':
			case '[object Character]': 
			case '[object Word]': 
			case '[object Line]': 
			case '[object Paragraph]': 
			case '[object TextColumn]':
					AddToArrayUnique(inputObject.parentStory,c_objects2Work);
					break;
			case '[object Group]':
					var i = inputObject.allPageItems.length;
						while (i--) 	CalcObjectForHandle(inputObject.allPageItems[i]);
					break;
			case '[object Rectangle]':
			case '[object GraphicLine]':
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
  
    if ( inpStory.leading == Leading.auto ) 
    { newLeadingSize = Leading.auto ; 
      } else  {  
    newLeadingSize = calcNewSize(inpStory.leading, percent);
    }

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
        while (c_selCount--) CalcObjectForHandle(c_currSelect[c_selCount]);
        var WorkLen = c_objects2Work.length;
        while (WorkLen--) HandleStory(c_objects2Work[WorkLen]);
      } catch (err) {
          alert (err); 
          exit(); 
      }
}

/** Возможность инклюда этого скрипта.
    Для того, что бы не выполнялась головная функция необходимо объявить переменную runCompactTXT и установить ей значение false */
try { runCompactTXT == undefined; } catch(err) { runCompactTXT = true; }
runCompactTXT &&  main () ;