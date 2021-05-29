const ps = new PerfectScrollbar('#cells', {
    wheelSpeed : 15,
    wheelPropagation: true
});

//cut copy paste
//scroll - updown
//formula

for(let i=1; i<100; ++i)
{
    let str = "";
    let n = i;
    while(n>0)
    {
        let rem = n%26;
        if(rem==0){
            str = "Z" + str;
            n = Math.floor(n/26) - 1;
        }
        else{
            str = String.fromCharCode( (rem-1)+65 ) + str;
            n =Math.floor(n/26);
        }
    }
    $('#columns').append(`<div class = "column-name column-${i}" id="${str}">${str}</div>`);
    $("#rows").append( `<div class="row-name">${i}</div>`);
}

let cellData = {
    "Sheet1" : {}
};

let selectedSheet = 'Sheet1';
let totalSheets = 1;
let lastlyAddedSheet = 1;
let saved = true;

let defaultProperties =  {
    "font-family" : 'Noto Sans',
    "font-size" : 14,
    "text" : "",
    "bold" : false,
    "underline" : false,
    "italic" : false,
    "alignment" : 'left',
    "color" : "#444",
    "bgcolor" : "#fff",
    "formula" : "",
    "upstream": [],
    "downStream" : []
};

for(let i=1; i<=100; i++)
{
    let row = $('<div class="cell-row" ></div>')
    for(let j=1; j<=100; ++j)
    {
        row.append( `<div id="row-${i}-col-${j}" class='input-cell' contenteditable='false'></div>` );
    }
    $("#cells").append(row);
}

$("#cells").scroll(function(e){
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);   
} )

$(".input-cell").dblclick(function(e){
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
    $(this).addClass('selected');
    $(this).attr("contenteditable", "true");
    $(this).focus();
})

$(".input-cell").blur( function(e){
    $(this).attr("contenteditable", "false");

    let [row, col] = getRowCol(this);
    if( cellData[selectedSheet][row-1] && cellData[selectedSheet][row-1][col-1] )
    {
        if( cellData[selectedSheet][row-1][col-1].formula != '' )
            updateStreams(this, []);
        
        cellData[selectedSheet][row-1][col-1].formula = '';
        
        updateCellData("text", $(this).text());
        let selfColCode = $(`.column-${col}`).attr('id');
        evalFormula(selfColCode+row);
    }
    // console.log(cellData);
})

function getRowCol(ele){
    // console.log( $(ele) );
    let id = $(ele).attr('id');
    let idArray = id.split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);

    return [rowId, colId];
}

function getTopLeftBottomRightCell(rowId, colId)
{
    let topCell = $( `#row-${rowId - 1}-col-${colId}` );
    let leftCell = $( `#row-${rowId}-col-${colId-1}` );
    let bottomCell = $( `#row-${rowId + 1}-col-${colId}` );
    let rightCell = $( `#row-${rowId}-col-${colId+1}` );

    return [ topCell, bottomCell, leftCell, rightCell];
}

$(".input-cell").click(function(e){
    // console.log(this);
    let [rowId, colId] = getRowCol(this);
    let [ topCell, bottomCell, leftCell, rightCell] = getTopLeftBottomRightCell(rowId, colId);

    if( $(this).hasClass('selected') && e.ctrlKey )
    {
        unselectCell( this, e, topCell, bottomCell, leftCell, rightCell );
    }
    else
        selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
})

function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell)
{
    if($(ele).attr('contenteditable') == 'false')
    {
        if($(ele).hasClass('top-selected')){
            topCell.removeClass('bottom-selected');
        }
    
        if($(ele).hasClass('bottom-selected')){
            bottomCell.removeClass('top-selected');
        }
    
        if($(ele).hasClass('right-selected')){
            rightCell.removeClass('left-selected');
        }
    
        if($(ele).hasClass('left-selected')){
            leftCell.removeClass('right-selected');
        }
    
        $(ele).removeClass("selected top-selected bottom-selected right-selected left-selected");
    }   

}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell)
{
    if(e.ctrlKey)
    {
        //top selected or not
        let topSelected;
        if(topCell){
            topSelected = topCell.hasClass("selected");
        }
        
        let bottomSelected;
        if(bottomCell){
            bottomSelected = bottomCell.hasClass("selected");
        }

        let rightSelected;
        if(rightCell){
            rightSelected = rightCell.hasClass("selected");
        }

        let leftSelected;
        if(leftCell){
            leftSelected = leftCell.hasClass("selected");
        }
        
        if(topSelected){
            $(ele).addClass('top-selected');
            topCell.addClass('bottom-selected')
        }

        if(bottomSelected){
            $(ele).addClass('bottom-selected');
            bottomCell.addClass('top-selected')
        }

        if(rightSelected){
            $(ele).addClass('right-selected');
            rightCell.addClass('left-selected')
        }

        if(leftSelected){
            $(ele).addClass('left-selected');
            leftCell.addClass('right-selected')
        }

    }
    else
    {
        $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");   
    }
    $(ele).addClass("selected");
    changeHeader( getRowCol(ele) );
}

function changeHeader([rowId, colId]){
    let data;
    if( cellData[selectedSheet][rowId-1] && cellData[selectedSheet][rowId-1][colId-1] )
    {
        data = cellData[selectedSheet][rowId-1][colId-1];
    }else
    {
        data = defaultProperties;
    }
    
    $('.alignment.selected').removeClass('selected');
    $(`.alignment[data-type = ${data.alignment}]`).addClass("selected");

    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, 'underline');

    $('#font-family').val(data['font-family']);
    $('#font-size').val(data['font-size']);
    $('#font-family').css( 'font-family', data['font-family'] );
    $("#fill-color").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color").css("border-bottom", `4px solid ${data.color}`);
    $('#formula-input').text(data.formula);
}

function addRemoveSelectFromFontStyle(data, property){
    if(data[property])
        $(`#${property}`).addClass('selected')
    else
        $(`#${property}`).removeClass('selected')
}

let startcellSelected = false;
let startCell={}, endCell={};
let scrollXRStarted = false;
let scrollXLStarted = false;
$(".input-cell").mousemove(function(e){
    e.preventDefault();
    if(e.buttons ==1 )
    {
        if( e.pageX > $(window).width()-10 && !scrollXRStarted )
        {
            scrollXR();
        }else if( e.pageX <(10) && !scrollXLStarted ){
            scrollXL();
        }

        if(!startcellSelected){
            let [rowId, colId] = getRowCol(this);
            startCell = { "rowId":rowId, "colId":colId };
            startcellSelected = true;
            selectAllBetweenCells(startCell, startCell);
        }
    }
    else{
        startcellSelected = false;
    }
});

$(".input-cell").mouseenter(function(e){
    if(e.buttons == 1){

        if(e.pageX <( $(window).width()-10 ) && scrollXRStarted ){
            clearInterval(scrollXRInterval);
            scrollXRStarted = false;
        }

        if(e.pageX > (10) && scrollXLStarted ){
            clearInterval(scrollXLInterval);
            scrollXLStarted = false;
        }   
        let [rowId, colId] = getRowCol(this);
        endCell = { "rowId":rowId, "colId":colId };
        selectAllBetweenCells(startCell, endCell);
    }
})

function selectAllBetweenCells(start, end){
    $(".input-cell.selected").removeClass("selected top-selected bottom-selected right-selected left-selected");
    for( let i = Math.min( start.rowId, end.rowId ); i<=Math.max(start.rowId, end.rowId); ++i )
    {
        for( let j = Math.min( start.colId, end.colId ); j<=Math.max( start.colId, end.colId ); ++j )
        {
            let [ topCell, bottomCell, leftCell, rightCell] = getTopLeftBottomRightCell(i, j);
            selectCell( $( `#row-${i}-col-${j}`)[0], {ctrlKey:true}, topCell, bottomCell, leftCell, rightCell);
        }
    }
}

let scrollXRInterval;
let scrollXLInterval;
function scrollXR(){
    scrollXRStarted = true;
    scrollXRInterval = setInterval(() => {
       $('#cells').scrollLeft( $("#cells").scrollLeft() + 100 ); 
    }, 100);
}

function scrollXL(){
    scrollXLStarted = true;
    scrollXLInterval = setInterval(() => {
       $('#cells').scrollLeft( $("#cells").scrollLeft() - 100 ); 
    }, 100);
}

$(".data-container").mousemove(function(e){
       e.preventDefault();
    if(e.buttons ==1 )
    {
        if( e.pageX > $(window).width()-10 && !scrollXRStarted )
        {
            scrollXR();
        }else if( e.pageX <(10) && !scrollXLStarted ){
            scrollXL();
        }
    }
})

$(".data-container").mouseup(function(e){
    clearInterval(scrollXRInterval);
    clearInterval(scrollXLInterval);
    scrollXLStarted = false;
    scrollXRStarted = false;
})

$(".alignment").click(function(e){
    let alignment = $(this).attr("data-type");
    $('.alignment.selected').removeClass('selected');
    $(this).addClass('selected');
    $('.input-cell.selected').css('text-align', alignment);
    // $('.input-cell.selected').each( function( index, data ){
    //     let [row, col] = getRowCol(data);
    //     cellData[row-1][col-1].alignment = alignment;
    // } )

    updateCellData("alignment", alignment);
})

$("#bold").click(function(e){
    setStyle(this, 'bold', 'font-weight');
})

$('#italic').click(function(e){
    setStyle( this, 'italic', 'font-style');
})

$('#underline').click(function(e){
    setStyle( this, 'underline', 'text-decoration');
})

function setStyle(ele, property, cssProp){
    if( $(ele).hasClass('selected') )
    {
        $(ele).removeClass('selected');
        $('.input-cell.selected').css( `${cssProp}` , '');
        // $('.input-cell.selected').each( function( index, data ){
        //     let [row, col] = getRowCol(data);
        //     cellData[row-1][col-1][property] = false;
        //     changeHeader([row, col]);
        // } )
        updateCellData( property, false );
    }
    else
    {
        $(ele).addClass('selected');
        $('.input-cell.selected').css(`${cssProp}`, property);
        // $('.input-cell.selected').each( function( index, data ){
        //     let [row, col] = getRowCol(data);
        //     cellData[row-1][col-1][property] = true;
        //     console.log( cellData[row-1][col-1] );
        //     changeHeader([row, col]);
        // } )
        updateCellData(property, true);
    }
}

$('.menu-selector').change(function(e){

    let value = $(this).val();
    let key = $(this).attr('id');
    if(key=='font-family'){
        $('#font-family').css(key, value);
    }

    if( !isNaN(value) ){
        value = parseInt(value);
    }

    $( '.input-cell.selected' ).css( key, value );
    // $( '.input-cell.selected' ).each( (index, data) => {
    //     let [row, col] = getRowCol(data);
    //     cellData[row-1][col-1][key] = value;
    // })

    updateCellData( key, value );
})

$(".pick-color").colorPick({
  'initialColor': '#abcd',
  'allowRecent': true,
  'recentMax': 5,
  'allowCustomColor': true,
  'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
  'onColorSelected': function(){
        if(this.color != '#ABCD'){
            if( $(this.element.children()[1]).attr('id') == 'fill-color' ){
                $('.input-cell.selected').css( 'background-color', this.color );
                $('#fill-color').css("border-bottom", `4px solid ${this.color}`)
                updateCellData("bgcolor", this.color);
            }else{
                $('.input-cell.selected').css( 'color', this.color );
                $('#text-color').css("border-bottom", `4px solid ${this.color}`)
                updateCellData('color', this.color);
            }
        }

    }
});

$('#fill-color, #text-color').click( function(){
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
} )

function updateCellData( property, value){
    let currentCellData = JSON.stringify(cellData);

    if(value != defaultProperties[property] ){
        $('.input-cell.selected').each( function(index, data){
            let [row, col] = getRowCol(data);
            if( cellData[selectedSheet][row-1] == undefined ){
                cellData[selectedSheet][row-1] = {};
                cellData[selectedSheet][row-1][col-1] = {...defaultProperties, "upStream":[], "downStream":[]};
                cellData[selectedSheet][row-1][col-1][property] = value;
            }
            else{
                if( cellData[selectedSheet][row-1][col-1] == undefined ){
                    // cellData[selectedSheet][row-1][col-1] = {};
                    cellData[selectedSheet][row-1][col-1] = {...defaultProperties, "upStream":[], "downStream":[]};
                    cellData[selectedSheet][row-1][col-1][property] = value;
                }
                else{
                    cellData[selectedSheet][row-1][col-1][property] = value;
                }
            }
        } );
    }
    else{
        $('.input-cell.selected').each( function(index, data){
            let [row, col] = getRowCol(data);
            if( cellData[selectedSheet][row-1] && cellData[selectedSheet][row-1][col-1] ){
                cellData[selectedSheet][row-1][col-1][property] = value;
                if( JSON.stringify(cellData[selectedSheet][row-1][col-1]) == JSON.stringify(defaultProperties) ){
                    delete cellData[selectedSheet][row-1][col-1];
                    if(Object.keys(cellData[selectedSheet][row-1]).length ==0 ){
                        delete cellData[selectedSheet][row-1];
                    }
                }
            }
        } );
    }

    if( saved && currentCellData != JSON.stringify(cellData))
        saved = false;
}

$('.container').click(function(e){
    $('.sheet-options-modal').remove();
})

function addSheetEvents(){

    $(".sheet-tab.selected").on("contextmenu", function(e){
        e.preventDefault();
        
        selectSheet(this);

        $('.sheet-options-modal').remove();
        let modal = $(`<div class='sheet-options-modal' >
                        <div class='option sheet-rename' > Rename</div>
                        <div class='option sheet-delete' > Delete</div>
                    </div>`)
    
        modal.css({'left': e.pageX} );
        $('.container').append(modal);
    
        $('.sheet-rename').click(function(e){
            let renameModal = $(`<div class='sheet-modal-parent' >
                                    <div class='sheet-rename-modal' >
                                        <div class='sheet-modal-title'> Rename Sheet </div>
                                        <div class='sheet-modal-input-container' >
                                            <span class='sheet-modal-input-title' >Rename Sheet to : </span>
                                            <input class='sheet-modal-input' type='text' />
                                        </div>
                                        <div class='sheet-modal-confirmation' >
                                            <div class='button yes-button' >OK</div>
                                            <div class='button no-button' > Cancel </div>
                                        </div>
                                    </div>
                                </div>`);
                                
            $('.container').append(renameModal);
            $('.sheet-modal-input').focus();
            $('.no-button').click(function(e){
                $('.sheet-modal-parent').remove();
            })
            $('.yes-button').click(function(e){
                renameSheet();
            })
            $('.sheet-modal-input').keypress(function(e){
                if(e.key=='Enter'){
                    renameSheet();
                }
            })
        })

        $(".sheet-delete").click(function(e){
            if(totalSheets>1){
                let deleteModal = $(`<div class='sheet-modal-parent' >
                                        <div class='sheet-delete-modal' >
                                            <div class='sheet-modal-title'> Sheet Name </div>
                                            <div class='sheet-modal-warning-container' >
                                                <span class='sheet-modal-warning-title' >Do you want to delete ?</span>
                                            </div>
                                            <div class='sheet-modal-confirmation' >
                                                <div class='button yes-button' >
                                                    <div class='material-icons delete-icon' >delete</div>
                                                    YES
                                                </div>
                                                <div class='button no-button' >NO</div>
                                            </div>
                                        </div>
                                    </div>`);
                
                $('.container').append(deleteModal);

                $('.no-button').click(function(e){
                    $('.sheet-modal-parent').remove();
                })
                
                $('.yes-button').click(function(e){
                    deleteSheet();
                })
            }else{
                alert('Not possible');
            }
        })

    });

    $(".sheet-tab.selected").click( function(e){
            selectSheet(this);
    } )
}

addSheetEvents();

$('.add-sheet').click( function(e){
    saved = false;
    ++lastlyAddedSheet;
    ++totalSheets;
    cellData[`Sheet${lastlyAddedSheet}`] = {};
    $('.sheet-tab.selected').removeClass('selected');
    $('.sheet-tab-container').append(`<div class='sheet-tab selected' >Sheet${lastlyAddedSheet}</div>`);
    selectSheet();
    addSheetEvents();
    $('.sheet-tab.selected')[0].scrollIntoView();
} )

function selectSheet(ele){
    if(ele && !$(ele).hasClass('selected')){
        $('.sheet-tab.selected').removeClass('selected');
        $(ele).addClass('selected');
    }
    emptyPreviousSheet();
    selectedSheet = $('.sheet-tab.selected').text();
    loadCurrentSheet();
    $('#row-1-col-1').click();
}

function emptyPreviousSheet(){
    
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys){
        let colKeys = Object.keys(data[i]);
        for( let j of colKeys){
            let cell = $(`#row-${parseInt(i)+1}-col-${parseInt(j)+1}`);
            cell.text("");
            cell.css( {
                "font-family" : 'Noto Sans',
                "font-size" : 14,
                "background-color" : '#fff',
                'color' : '#444',
                'font-weight' : '',
                'font-style': '',
                'text-decoration' : '',
                'text-align' : 'left'
            } );
        }
    }
}

function loadCurrentSheet(){
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys){
        let colKeys = Object.keys(data[i]);
        for( let j of colKeys){
            let cell = $(`#row-${parseInt(i)+1}-col-${parseInt(j)+1}`);
            cell.text( data[i][j].text );
            cell.css( {
                "font-family" : data[i][j]['font-family'],
                "font-size" : data[i][j]['font-size'],
                "background-color" : data[i][j]['bgcolor'],
                'color' : data[i][j].color,
                'font-weight' : data[i][j].bold ? "bold" : "",
                'font-style': data[i][j].italic ? "italic" : "",
                'text-decoration' : data[i][j].underline ? "underline" : "",
                'text-align' : data[i][j].alignment
            } );
        }
    }
}

function renameSheet(){
    let newSheetName = $('.sheet-modal-input').val();

    if( newSheetName && !Object.keys(cellData).includes(newSheetName) )
    {
        let newCellData = {};
        for(let i of Object.keys(cellData))
        {
            if( i!=selectedSheet )
            {
                newCellData[i] = cellData[i];
            }else{
                newCellData[newSheetName] = cellData[i];
            }
        }
        cellData = newCellData;
        saved = false;  
        selectedSheet = newSheetName;
        
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
    }
    else{
        $('.rename-error').remove();
        $('.sheet-modal-input-container').append(`
        <div class='rename-error'>Sheet name is not valid or sheet already exists !!</div>`);
    }
}

function deleteSheet(){
    let sheetIndex = Object.keys(cellData).indexOf(selectedSheet);
    let currSelectedSheet = $('.sheet-tab.selected');
    let nextSheet;
    if( sheetIndex==0 )
    {
        nextSheet = currSelectedSheet.next()[0];
    }else{
        nextSheet = currSelectedSheet.prev()[0];
    }
    currSelectedSheet.remove();
    selectSheet(nextSheet);
    delete cellData[currSelectedSheet.text()];
    $(".sheet-modal-parent").remove();
    --totalSheets;
}

$('.left-scroller, .right-scroller').click(function(e){
    console.log('helo')
    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if( selectedSheetIndex != 0 && $(this).text() =='arrow_left' )
    {
        selectSheet( $('.sheet-tab.selected').prev()[0] );
    }else if( selectedSheetIndex != (keysArray.length-1) && $(this).text()=='arrow_right' ){
        selectSheet( $('.sheet-tab.selected').next()[0] );
    }

    $('.sheet-tab.selected')[0].scrollIntoView();
})

$('#menu-file').click(function(){
    let fileModal = $(`<div class='file-modal' >
                            <div class='options-menu' >
                                <div class='close'>
                                    <div class='close-icon material-icons' >arrow_circle_down</div>
                                    <div>Close</div>
                                </div>
                                <div class='new' >
                                    <div class='new-icon material-icons' >insert_drive_file</div>
                                    <div>New</div>
                                </div>
                                <div class='open' >
                                    <div class='open-icon material-icons' >folder_open</div>
                                    <div>Open</div>
                                </div>
                                <div class='save' >
                                    <div class='save-icon material-icons' >save</div>
                                    <div>Save</div>
                                </div>
                            </div>
                            <div class='recent-files' ></div>
                            <div class='file-transparent' ></div>
                        </div>`);
    
    $('.container').append(fileModal);
    fileModal.animate({
        width : '100vw'
    }, 500);
    $('.close, .file-transparent, .save, .open').click( function(e){
        fileModal.remove();
    } )

    $('.new').click( function(e){

        if( saved ){
            newFile();
        }
        else{
            $('.container').append(` <div class='sheet-modal-parent' >
                                        <div class='sheet-delete-modal' >
                                            <div class='sheet-modal-title'> ${ $('.title').text()} </div>
                                            <div class='sheet-modal-warning-container' >
                                                <span class='sheet-modal-warning-title' >Do you want to save changes ?</span>
                                            </div>
                                            <div class='sheet-modal-confirmation' >
                                                <div class='button yes-button' >YES</div>
                                                <div class='button no-button' >NO</div>
                                            </div>
                                        </div>
                                    </div>`);
        
            $('.no-button').click(function(){
                $('.sheet-modal-parent').remove();
                newFile();
            })
            $('.yes-button').click(function(){
                $('.sheet-modal-parent').remove();
                saveFile(true);
                // newFile();
            })    
        }
    } )

    $('.save').click( function(){
        if(!saved)
            saveFile();
    })

    $('.open').click(function(){
        openFile();
    })

});

function newFile(){
    emptyPreviousSheet();
    cellData = {"Sheet1" : {}};
    $('.sheet-tab').remove();
    $( '.sheet-tab-container').append( `<div class='sheet-tab selected'>Sheet1</div>` );
    addSheetEvents();
    selectedSheet = 'Sheet1';
    totalSheets = 1;
    lastlyAddedSheet = 1;
    $('.title').text('Excell - Book');
    $('#row-1-col-1').click();
}

function saveFile(newClicked){
    $('.container').append(`<div class='sheet-modal-parent' >
                                <div class='sheet-rename-modal' >
                                    <div class='sheet-modal-title'> Save File </div>
                                    <div class='sheet-modal-input-container' >
                                        <span class='sheet-modal-input-title' >File Name : </span>
                                        <input class='sheet-modal-input' type='text' value='${$('.title').text()}' />
                                    </div>
                                    <div class='sheet-modal-confirmation' >
                                        <div class='button yes-button' >Save</div>
                                        <div class='button no-button' > Cancel </div>
                                    </div>
                                </div>
                            </div>`); 

    $('.yes-button').click(function(){
        let anchorTag = $(`<a href='data:application/json,${encodeURIComponent(JSON.stringify(cellData))}' download="${$('.title').text()+'.json'}"></a>`)
        $('.container').append(anchorTag);
        $(anchorTag)[0].click();
        $(anchorTag).remove();
        saved = true;
    })
    $('.no-button, .yes-button').click( function(e){
        $('.sheet-modal-parent').remove();
        if(newClicked){
            newFile();
        }
    } )                          
}

function openFile(){
    let inputFile = $(`<input type="file" />` );
    $('.container').append(inputFile);
    inputFile.click();
    console.log('hello');
    
    inputFile.change(function(e){
        console.log(e);

        let file = e.target.files[0];
        $('.title').text(file.name.split(".json")[0]);
        let reader = new FileReader();
        reader.readAsText(file);
        reader.onload = () => {
            cellData = JSON.parse(reader.result);
            $('.sheet-tab').remove();
            
            let sheets = Object.keys(cellData);
            for( let i of sheets){
                if( i.includes('Sheet') ){
                    let splittedSheetArray = i.split("Sheet");
                    if( splittedSheetArray.length==2 && !isNaN(splittedSheetArray[1]) ){
                        lastlyAddedSheet = parseInt(splittedSheetArray[1]);
                    }
                }
                $('.sheet-tab-container').append(`<div class='sheet-tab selected' >${i}</div>`);
            }
            addSheetEvents();
            $(".sheet-tab").removeClass('selected');
            $( $('.sheet-tab')[0] ).addClass('selected');
            selectedSheet = sheets[0];
            totalSheets = sheets.length;
            lastlyAddedSheet = totalSheets;
            loadCurrentSheet();
            inputFile.remove();        
        }
    })

}

let clipboard = { startCell:[], cellData:{} };
let isContentCut = false;
$(".copy, .cut").click(function(){
    clipboard = { startCell:[], cellData:{} }
    clipboard.startCell = getRowCol( $('.input-cell.selected')[0] );

    $(".input-cell.selected").each(function(index, data){
        let [row, col] = getRowCol(data);
        if( cellData[selectedSheet][row-1] && cellData[selectedSheet][row-1][col-1] ){
            if(!clipboard.cellData[row]){
                clipboard.cellData[row] = {};
            }
            clipboard.cellData[row][col] = {...cellData[selectedSheet][row-1][col-1]}
        }
    })

    if($(this).text() == 'content_cut')
    {
        isContentCut = true;
    }
} )

$('.paste').click(function(){

    if(isContentCut){
        emptyPreviousSheet();
    }

    let startCell = getRowCol( $('.input-cell.selected')[0] );
    for(let i of Object.keys(clipboard.cellData) )
    {
        let cols = Object.keys(clipboard.cellData[i]);
        for( let j of cols ){

            if(isContentCut){
                delete cellData[selectedSheet][i-1][j-1];
                if(Object.keys(cellData[selectedSheet][i-1]).length == 0 ){
                    delete cellData[selectedSheet][i-1];
                }
            }
        }
    }

    for( let i of Object.keys(clipboard.cellData))
    {
        let cols = Object.keys(clipboard.cellData[i]);
        for(let j of cols)
        {
            let rd = parseInt(i) - parseInt(clipboard.startCell[0]);
                let cd = parseInt(j) - parseInt(clipboard.startCell[1]);
                if(!cellData[selectedSheet][startCell[0] + rd-1]){
                    cellData[selectedSheet][startCell[0] + rd-1] = {};
                }
                cellData[selectedSheet][startCell[0] + rd-1][startCell[1]+cd-1] = {...clipboard.cellData[i][j]};
        }
    }

    loadCurrentSheet();
    if(isContentCut){
        isContentCut=false;
        clipboard = { startCell:[], cellData:{} };
    }
    
})

$("#formula-input").blur(function(e){
    if( $(".input-cell.selected").length>0 ){
        let formula = $(this).text();
        let tempElements = formula.split(" ");
        let elements = [];
        for(let i of tempElements){
            if(i.length >= 2){
                i = i.replace( "(","");
                i = i.replace( ")","");
                if( !elements.includes(i) )
                    elements.push(i);
            }
        }
        $('.input-cell.selected').each( function(index, data){
            
            if(updateStreams(data, elements, false)){
                let [row, col] = getRowCol(data);
                cellData[selectedSheet][row-1][col-1].formula = formula;
                let selfColCode = $(`.column-${col}`).attr('id');
                evalFormula(selfColCode + row);
            }else{
                alert("Formula invalid");
            }
        } )
        
    }else{
        alert("Please select a cell !!");
    }
})

function updateStreams(ele, elements, update, oldUpstream){
    let [row, col] = getRowCol(ele);

    let selfColCode = $(`.column-${col}`).attr("id");
    if(elements.includes(selfColCode + row)){
        return false;
    }

    if(cellData[selectedSheet][row-1] && cellData[selectedSheet][row-1][col-1] ){
        let downStream = cellData[selectedSheet][row-1][col-1].downStream;
        let upStream = cellData[selectedSheet][row-1][col-1].upStream;
        
        for(let i of downStream){
            if(elements.includes(i))
                return false;
        }
        
        for(let i of downStream){
            let [calRow, calCol] = codeToValue(i);
            console.log( updateStreams( $(`#row-${calRow}-col-${calCol}`)[0], elements, true, upStream ) )
        }
    }

    if( !cellData[selectedSheet][row-1] ){
        cellData[selectedSheet][row-1] = {};  
        cellData[selectedSheet][row-1][col-1] = { ...defaultProperties, "upStream":[...elements], "downStream":[] };  
    }
    else if( !cellData[selectedSheet][row-1][col-1] ){
        cellData[selectedSheet][row-1][col-1] = { ...defaultProperties, "upStream":[...elements], "downStream":[] };
    }
    else{
        let upStream = [...cellData[selectedSheet][row-1][col-1].upStream];
        if(update)
        {
            for(let i of oldUpStream){
                let [calRow, calCol] = codeToValue(i);
                let index = cellData[selectedSheet][calRow-1][calCol-1].downStream.indexOf(selfColCode+row)
                cellData[selectedSheet][calRow-1][calCol-1].downStream.splice(index, 1);
                if( JSON.stringify(cellData[selectedSheet][calRow-1][calCol-1]) == JSON.stringify(defaultProperties) ){
                    delete cellData[selectedSheet][calRow-1][calCol-1];
                    if( Object.keys(cellData[selectedSheet][calRow-1]).length==0 )
                        delete cellData[selectedSheet][calRow-1];
                }
                index = cellData[selectedSheet][row-1][col-1].upStream.indexOf(i);
                cellData[selectedSheet][row-1][col-1].upStream.splice(index, i);
            }
            for(let i of elements){
                cellData[selectedSheet][row-1][col-1].upStream.push(i);
            }
        }
        else{
            for( let i of upStream){
                let [calRow, calCol] = codeToValue(i);
                let index = cellData[selectedSheet][calRow-1][calCol-1].downStream.indexof(selfColCode+row);
                cellData[selectedSheet][calRow-1][calCol-1].downStream.splice(index, 1);
                if(JSON.stringify(cellData[selectedSheet][calRow-1][calCol-1]) == JSON.stringify(defaultProperties) ){
                    delete cellData[selectedSheet][calRow-1][calCol-1];
                    if(Object.keys(cellData[selectedSheet][calRow-1][calCol-1]).length == 0)
                        delete cellData[selectedSheet][calRow]-1;
                }
            }
            cellData[selectedSheet][row-1][col-1].upStream = [...elements];
        }
    }
    
    for(let i of elements){
        let [calRow, calCol] = codeToValue(i);

        if( !cellData[selectedSheet][calRow-1] ){
            cellData[selectedSheet][calRow-1] = {};  
            cellData[selectedSheet][calRow-1][calCol-1] = { ...defaultProperties, "upStream":[...elements], "downStream":[selfColCode+row] };  
        }
        else if( !cellData[selectedSheet][calRow-1][calCol-1] ){
            cellData[selectedSheet][calRow-1][calCol-1] = { ...defaultProperties, "upStream":[...elements], "downStream":[selfColCode+row] };
        }
        else{
            cellData[selectedSheet][calRow-1][calCol-1].downStream.push(selfColCode + row);
        }
    }

    return true;
}

function codeToValue(code){
    let colCode = "";
    let rowCode = "";
    for(let i=0; i<code.length; ++i){
        if( !isNaN(code.charAt(i) ) ){
            rowCode += code.charAt(i);
        }
        else{
            colCode += code.charAt(i);
        }
    }
    let colId = parseInt( $(`#${colCode}`).attr('class').split(" ")[1].split("-")[1] );
    let rowId = parseInt(rowCode);

    return [rowId, colId];
}

function evalFormula(cell){
    
    let [row, col] = codeToValue(cell);
    let formula = cellData[selectedSheet][row-1][col-1].formula;

    if(formula!= "")
    {
        let upStream = cellData[selectedSheet][row-1][col-1].upStream;
        let upStreamValue = [];
        for(let i in upStream)
        {
            let [calRow, calCol] = codeToValue(upStream[i]);
            let value;
            if( cellData[selectedSheet][calRow-1][calCol-1].text == "" )
                value='0';
            else
                value = cellData[selectedSheet][calRow-1][calCol-1].text;
            
            upStreamValue.push( value );
            formula = formula.replace( upStream[i], upStreamValue[i] );
        }
        cellData[selectedSheet][row-1][col-1].text = eval(formula);
        loadCurrentSheet();
    }

    let downStream = cellData[selectedSheet][row-1][col-1].downStream;
    for( let i=downStream.length-1; i>=0; i--)
    {
        evalFormula(downStream[i]);
    }
}