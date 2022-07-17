var currentWindow = window.self;

//let myWindow_width = 900;
let myWindow_width = screen.width;
let myWindow_heigth = 250;
//var myWindow_left = (screen.width/2)-(myWindow_width/2);
var myWindow_left = screen.left;
//var myWindow_top = (screen.height/2)-(myWindow_heigth/2);
//var myWindow_top = (screen.height)-(myWindow_heigth);
var myWindow_top = 0;

var myWindow = window.open('', '_blank', 'height='+myWindow_heigth+',width='+myWindow_width+',top='+myWindow_top+',left='+myWindow_left+',scrollbars=yes,status=no,location=no,toolbar=no,menubar=no');




var myScript = myWindow.document.createElement("script");
myScript.innerHTML = '(' + string_SpreadsheetMethods + ')(window, window.document);';
myWindow.document.head.appendChild(myScript);



const spreadsheetDiv = myWindow.document.createElement("div");
spreadsheetDiv.setAttribute("id", "mySpreadsheetContainer");
myWindow.document.body.appendChild(spreadsheetDiv);





var result_array = [];

var myUl = document.getElementById("results");
var list = myUl.getElementsByTagName("li");
for (var i = 0; i < list.length; i++) 
{
  let extracted_array = [];




  var list_img = list[i].querySelectorAll("img");
  var img_src = '';
  if(typeof list_img !== 'undefined')
  {
    if(typeof list_img[0] !== 'undefined')
    {
        if(list_img[0].hasAttribute('src'))
        {
          img_src = list_img[0].getAttribute('src')
        }
    }
    else
    {
      list_img = list[i].getElementsByTagName("noscript")[0];
      var nosHtml = list_img.textContent||list_img.innerHTML; 

      if ( nosHtml )
      {
        var temp = document.createElement("div");
        temp.innerHTML = nosHtml;
      
        
        document.body.appendChild(temp);
        //var ex = document.getElementsByTagName("img")[0];
        document.body.removeChild(temp);

        img_src = temp.getElementsByTagName("img")[0].getAttribute('style').replace('background-image:','');
      }
    }

    if(img_src != '')
    {
        extracted_array.push(String(img_src));
    }
    else{
      extracted_array.push('ندارد');
    }
  }
  else{
      extracted_array.push('ندارد');
  }
  

  
  
  var list_header = list[i].querySelectorAll("h2")[0].innerHTML;
  if(typeof list_header !== 'undefined')
  {
    extracted_array.push(String(list_header));
  }
  else
  {
    extracted_array.push('ندارد');
  }



  var list_price = list[i].querySelectorAll("footer")[0].getElementsByClassName("_2NIRZ")[0].innerHTML.replaceAll('<!-- -->','');
  if(typeof list_price !== 'undefined')
  {
    extracted_array.push(String(list_price));
  }
  else
  {
    extracted_array.push('ندارد');
  }


  
  var list_score = list[i].querySelectorAll("footer")[0].getElementsByClassName("_1kN9Z")[0];
  if(typeof list_score !== 'undefined')
  {
    extracted_array.push(String(list_score.innerHTML));
  }
  else
  {
    extracted_array.push('ندارد');
  }



  var list_url = list[i].querySelectorAll('[data-houseid]')[0];
  if(typeof list_url !== 'undefined')
  {
    extracted_array.push(String(list_url.getAttribute("href")));
  }
  else
  {
    extracted_array.push('ندارد');
  }
  


  var list_id = list[i].querySelectorAll('[data-houseid]')[0];
  if(typeof list_id !== 'undefined')
  {
    extracted_array.push(String(list_id.getAttribute("data-houseid")));
  }
  else
  {
    extracted_array.push('ندارد');
  }


  
  var list_id = list[i].querySelectorAll('[data-houseid]')[0];
  if(typeof list_id !== 'undefined')
  {
    extracted_array.push(String(list_id.getAttribute("data-houseid")));
  }
  else
  {
    extracted_array.push('ندارد');
  }
  

  

  result_array.push(extracted_array);
  
}
var serialize_result_array = JSON.stringify(result_array);



myScript = myWindow.document.createElement("script");
myScript.innerHTML = string_createOwnSpreadsheet;
myScript.innerHTML += 'setTimeout(createOwnSpreadsheet('+serialize_result_array+'),1000);';
myWindow.document.head.appendChild(myScript);








var string_createOwnSpreadsheet = function createOwnSpreadsheet(result_array) {

  const mySpreadsheet = Spreadsheet('#mySpreadsheetContainer');

    mySpreadsheet.createSpreadsheet(
      {
        'عکس': 'img',
        'مشخصات': 'text',
        'قیمت': 'text',
        'امتیاز': 'text',
        'لینک': 'link',
        'نمایش در صفحه': 'button',
        'عملیات': 'action button'
      },
      {
        rowCount: result_array.length,
        persistent: true,
        //submitCallback: (tableArray) => console.log(JSON.stringify(tableArray)),
        data: result_array
      }
    );
    
}





function insertAfter(newNode, referenceNode) {
  
    referenceNode.parentNode.insertBefore(newNode, referenceNode);
}


/*function my_replaceChild(newNode, oldNode) {

    newNode.parentNode.replaceChild(oldNode, newNode);
}*/





var string_SpreadsheetMethods = function (global, document) {
  function Spreadsheet(selector) {


    
    var myStyles = 'tr:hover{background-color: #ddd;}';
    var style = document.createElement('style');
    if (style.styleSheet) 
    {
        style.styleSheet.cssText = myStyles;
    } 
    else 
    {
        style.appendChild(document.createTextNode(myStyles));
    }
    document.getElementsByTagName('head')[0].appendChild(style);


    
    
    const _self = {};
  
    const container = document.querySelector(selector);

    const styles = {
      border: '1px solid grey',
    };

    container.style.padding = '10px';

    const getTable = () => container.firstElementChild.firstElementChild;

    _self.createSpreadsheet = function (
      columns,
      options = { rowCount: 1, persistent: false, data: [], submitCallback }
    ) {

      const columnNames = Object.keys(columns);
      const tableContainer = document.createElement('div');
      container.appendChild(tableContainer);

      const spreadsheetTable = document.createElement('table');
      tableContainer.appendChild(spreadsheetTable);
      spreadsheetTable.style.border = styles.border;
      spreadsheetTable.style.fontFamily = 'Arial, sans-serif';
      spreadsheetTable.style.borderCollapse = 'collapse';
      spreadsheetTable.style.borderRadius = '100px';
      spreadsheetTable.style.width = '100%';

      const headerRow = document.createElement('tr');
      spreadsheetTable.appendChild(headerRow);

      for (const columnName of columnNames) {
        const headerRowCell = document.createElement('th');
        headerRow.appendChild(headerRowCell);
        headerRowCell.append(columnName);
        headerRowCell.style.border = styles.border;
        headerRowCell.style.borderCollapse = 'collapse';
        headerRowCell.style.borderRadius = '100px';
        headerRowCell.style = "border: 1px solid grey;border-collapse: collapse;padding-top: 12px;padding-bottom: 12px;background-color: #75c1cb;color: white;";
      }

      const buttonContainer = document.createElement('div');
      container.appendChild(buttonContainer);
      buttonContainer.style.marginTop = '10px';

      const addRowButton = createButton(() => addRows(columns), 'Add row');
      const deleteBottomRowButton = createButton(
        () => deleteBottomRow(),
        'Delete bottom row'
      );
      const clearButton = createButton(
        () => clearSpreadsheet(options.persistent !== undefined),
        'Clear sheet'
      );

      //buttonContainer.append(addRowButton, deleteBottomRowButton, clearButton);

      if (options.submitCallback) {
        const submitButton = createButton(
          () => options.submitCallback(this.arrayify()),
          'Submit'
        );
        buttonContainer.appendChild(submitButton);
      }

      if (options.persistent) {
        const saveSpreadsheetButton = createButton(
          () => saveSpreadsheet(),
          'Save'
        );
        buttonContainer.appendChild(saveSpreadsheetButton);
      }

      loadSpreadsheet(columns, options);
    };

    _self.addCellStyle = function (value, style) {
      const table = getTable();

      for (const cell of table.getElementsByTagName('td')) {
        const cellInput = cell.firstElementChild.firstElementChild;
        if (cellInput.value === value) {
          cell.style.borderRadius = '0px';
          cell.style.background = style;
          cellInput.style.background = style;
        }
      }
    };

    _self.addCellStyleSheet = function (stylingMap) {
      const table = getTable();

      for (const cell of table.getElementsByTagName('td')) {
        const cellInput = cell.firstElementChild.firstElementChild;
        const value = cellInput.value;
        if (value in stylingMap) {
          cell.style.borderRadius = '0px';
          cell.style.background = stylingMap[value];
          cellInput.style.background = stylingMap[value];
        }
      }
    };

    _self.arrayify = function () {
      const table = getTable();
      const tableArray = [];

      [...table.children].forEach((row, i) => {
        if (i !== 0) {
          const rowArray = [];
          [...row.children].forEach((cell) => {

            let my_elementName = cell.firstElementChild.firstElementChild.tagName.toLowerCase();
            if(my_elementName == 'img')
            {
              const cellValue = cell.firstElementChild.firstElementChild.getAttribute('src');
              rowArray.push(cellValue);
            }
            else if(my_elementName == 'a')
            {
              const cellValue = cell.firstElementChild.firstElementChild.getAttribute('href');
              rowArray.push(cellValue);
            }
            else if(my_elementName == 'button')
            {
              const cellValue = cell.firstElementChild.firstElementChild.getAttribute('uniqueId');
              rowArray.push(cellValue);
            }
            else
            {
              const cellValue = cell.firstElementChild.firstElementChild.value;
              rowArray.push(cellValue);
            }
            
          });
          tableArray.push(rowArray);
        }
      });
      
      return tableArray;
    };

    const createButton = function (onclick, innerText) {
      const button = document.createElement('button');
      button.onclick = onclick;
      button.innerText = innerText;
      button.style.marginLeft = '5px';
      button.style.marginRight = '5px';

      return button;
    };

    const addRows = function (columnTypes, rowCount = 1) {
      const table = getTable();

      for (let i = 0; i < rowCount; i++) {

        let j = 0;

        
        const cellRow = document.createElement('tr');
        //cellRow.setAttribute('id', column_copy[i][j]);
        
        for (const columnType in columnTypes) {
          
          const cell = document.createElement('td');
          cellRow.appendChild(cell);
          cell.style.border = styles.border;
          cell.style.borderCollapse = 'collapse';
          //cell.style.borderRadius = '100px';
          cell.style.fontFamily = 'Tahoma';

          const cellDiv = document.createElement('div');
          cell.appendChild(cellDiv);
          //cellDiv.style.resize = 'horizontal';
          cellDiv.style.overflow = 'auto';
          cellDiv.style.textAlign = 'center';
          cellDiv.style.width = '100%';

          if(columnTypes[columnType]=='img')
          {
              const cellInput = document.createElement('img');
              cellDiv.appendChild(cellInput);
              cellInput.style.width = '100px';
              cellInput.style.height = '100px';
              cellInput.style.border = 'none';
              cellInput.setAttribute('src', column_copy[i][j]);
              cellInput.style.textAlign = 'center';
          }
          else if(columnTypes[columnType]=='link')
          {
              const cellInput = document.createElement('a');
              cellDiv.appendChild(cellInput);
              cellInput.innerHTML = 'باز کردن';
              cellInput.setAttribute('href', column_copy[i][j]);
              cellInput.setAttribute('onclick', 'window.opener.location.href = window.opener.location.protocol + "//" + window.opener.location.hostname + "'+column_copy[i][j]+'";return false;');
              cellInput.style.textAlign = 'center';
              cellInput.style = 'text-align: center;color: white;background-color: #1ebba3;display: inline-block;padding: 10px 20px;border: 2px solid #099983;text-decoration: none;text-align: center;';
          }
          else if(columnTypes[columnType]=='button')
          {
              const cellInput = document.createElement('button');
              cellDiv.appendChild(cellInput);
              cellInput.innerHTML = 'نمایش';
              cellInput.setAttribute('uniqueId', column_copy[i][j]);
              cellInput.setAttribute('onclick', 'window.opener.document.querySelectorAll(`[data-houseid="'+column_copy[i][j]+'"]`)[0].parentNode.closest("li").style.border = "10px solid red";return false;');
              cellInput.style = 'text-align: center;padding: 10px;cursor: pointer;font-family: inherit;font-size: 15px;border-radius: 10px;background-color: lightcyan;';
          }
          else if(columnTypes[columnType]=='action button')
          {
              const cellInput = document.createElement('button');
              cellDiv.appendChild(cellInput);
              cellInput.innerHTML = 'قبلی';
              cellInput.setAttribute('class', 'action-button');
              cellInput.setAttribute('uniqueId', column_copy[i][j]);
              cellInput.setAttribute('onclick', 'if(this.parentNode.closest("tr").parentNode.getElementsByTagName("tr")[0] == this.parentNode.closest("tr").previousSibling){return false;};this.parentNode.closest("tr").previousSibling.parentNode.insertBefore(this.parentNode.closest("tr"),this.parentNode.closest("tr").previousSibling);window.opener.insertAfter(window.opener.document.querySelectorAll(`[data-houseid="'+column_copy[i][j]+'"]`)[0].parentNode.closest("li"),window.opener.document.querySelectorAll(`[data-houseid="'+column_copy[i][j]+'"]`)[0].parentNode.closest("li").previousSibling);');
              cellInput.style = 'text-align: center;padding: 10px;cursor: pointer;font-family: inherit;font-size: 15px;border-radius: 50%;background-color: #ffd2d2;';


              const cellInput2 = document.createElement('button');
              cellDiv.appendChild(cellInput2);
              cellInput2.innerHTML = 'بعدی';
              cellInput2.setAttribute('class', 'action-button');
              cellInput2.setAttribute('uniqueId', column_copy[i][j]);
              cellInput2.setAttribute('onclick', 'this.parentNode.closest("tr").parentNode.insertBefore(this.parentNode.closest("tr").nextSibling,this.parentNode.closest("tr"));window.opener.insertAfter(window.opener.document.querySelectorAll(`[data-houseid="'+column_copy[i][j]+'"]`)[0].parentNode.closest("li").nextSibling,window.opener.document.querySelectorAll(`[data-houseid="'+column_copy[i][j]+'"]`)[0].parentNode.closest("li"));');
              cellInput2.style = 'text-align: center;padding: 10px;cursor: pointer;font-family: inherit;font-size: 15px;border-radius: 50%;background-color: #b5ffb8;margin-left: 3px;';
          }
          else
          {
              const cellInput = document.createElement('input');
              cellDiv.appendChild(cellInput);
              cellInput.setAttribute('type', columnTypes[columnType]);
              cellInput.style.width = '100%';
              cellInput.style.border = 'none';
              cellInput.style.textAlign = 'center';
              cellInput.style.fontFamily = 'inherit';
              cellInput.style.direction = 'rtl';
              cellInput.style.backgroundColor = 'transparent';
          }
          j++;
        }

        table.appendChild(cellRow);
      }
    };

    const deleteBottomRow = function () {
      const table = getTable();

      if (table.children.length > 2) {
        table.removeChild(table.lastElementChild);
      } else {
        alert('Cannot delete any more rows!');
      }
    };

    const saveSpreadsheet = function () {
      const tableArray = _self.arrayify();
      console.log(tableArray);
      localStorage.setItem(selector, JSON.stringify(tableArray));
    };
var column_copy = null;
    const loadSpreadsheet = function (
      columnTypes,
      { data = [], rowCount = 1, persistent = false }
    ) {
      const table = getTable();
      let tableArray;
      column_copy = data;
      
      if (data.length > 0) {
        tableArray = data;
      }

      if (!persistent) {
        localStorage.removeItem(selector);
      } else if (localStorage.getItem(selector)) {
        tableArray = JSON.parse(localStorage.getItem(selector));
      }

      const rowDifference = tableArray ? tableArray.length : rowCount;

      addRows(columnTypes, rowDifference);

      if (tableArray) {

        var myUl = window.opener.document.getElementById("results");
        var list = window.opener.myUl.getElementsByTagName("li");

        
        [...table.children].forEach((row, i) => {
          if (i !== 0) {
            
            [...row.children].forEach((cell, j) => {

              let my_elementName = cell.firstElementChild.firstElementChild.tagName.toLowerCase();
              if(my_elementName == 'img')
              {
                cell.firstElementChild.firstElementChild.setAttribute('src',tableArray[i - 1][j]);
              }
              else if(my_elementName == 'a')
              {
                cell.firstElementChild.firstElementChild.setAttribute('href',tableArray[i - 1][j]);
                cell.firstElementChild.firstElementChild.setAttribute('onclick', 'window.opener.location.href = window.opener.location.protocol + "//" + window.opener.location.hostname + "'+tableArray[i - 1][j]+'";return false;');
              }
              else if(my_elementName == 'button')
              {
                for (var k = 0; k < list.length;k++) 
                {
                  let selected_li = window.opener.document.querySelectorAll(`[data-houseid="`+tableArray[i - 1][j]+`"]`)[0];
                  
                  if(selected_li !== 'undefined')
                  {
                    if(selected_li.parentNode !== 'undefined')
                    {
                      if(selected_li.parentNode.closest("li") == list[k])
                      {
                        console.log(tableArray[i - 1][j]);
                        window.opener.insertAfter(list[k],list[i-1]);
                      }
                    }
                  }
                  
                }
                
                if(cell.firstElementChild.firstElementChild.classList.contains('action-button'))
                {
                  cell.firstElementChild.childNodes[0].setAttribute('uniqueId', tableArray[i - 1][j]);
                  cell.firstElementChild.childNodes[0].setAttribute('onclick', 'if(this.parentNode.closest("tr").parentNode.getElementsByTagName("tr")[0] == this.parentNode.closest("tr").previousSibling){return false;};this.parentNode.closest("tr").previousSibling.parentNode.insertBefore(this.parentNode.closest("tr"),this.parentNode.closest("tr").previousSibling);window.opener.insertAfter(window.opener.document.querySelectorAll(`[data-houseid="'+tableArray[i - 1][j]+'"]`)[0].parentNode.closest("li"),window.opener.document.querySelectorAll(`[data-houseid="'+tableArray[i - 1][j]+'"]`)[0].parentNode.closest("li").previousSibling);');

                  cell.firstElementChild.childNodes[1].setAttribute('uniqueId', tableArray[i - 1][j]);
                  cell.firstElementChild.childNodes[1].setAttribute('onclick', 'this.parentNode.closest("tr").parentNode.insertBefore(this.parentNode.closest("tr").nextSibling,this.parentNode.closest("tr"));window.opener.insertAfter(window.opener.document.querySelectorAll(`[data-houseid="'+tableArray[i - 1][j]+'"]`)[0].parentNode.closest("li").nextSibling,window.opener.document.querySelectorAll(`[data-houseid="'+tableArray[i - 1][j]+'"]`)[0].parentNode.closest("li"));');
                }
                else
                {
                  cell.firstElementChild.firstElementChild.setAttribute('uniqueId', tableArray[i - 1][j]);
                  cell.firstElementChild.firstElementChild.setAttribute('onclick', 'window.opener.document.querySelectorAll(`[data-houseid="'+tableArray[i - 1][j]+'"]`)[0].parentNode.closest("li").style.border = "10px solid red";return false;');
                }
                
              }
              else
              {
                cell.firstElementChild.firstElementChild.value =
                tableArray[i - 1][j];
              }
              
              
            });
          }
        });
      }
    };

    const clearSpreadsheet = function (save) {
      const table = getTable();
      [...table.children].forEach((row, i) => {
        if (i !== 0) {
          [...row.children].forEach((cell) => {
            cell.firstElementChild.firstElementChild.value = '';
          });
        }
      });
      if (save) {
        saveSpreadsheet();
      }
    };

    return _self;
  }

  global.Spreadsheet = global.Spreadsheet || Spreadsheet;
}








