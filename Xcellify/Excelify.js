/*jshint expr:true*/
/*
 * Excelify by Sam Larison
 * transform selection of input elements into Excel-like spreadsheet with undo functionality
 * features: spreadsheet copy, cut, paste, multi-select, undo, redo, clear multiple, enter key move to next cell in selection
*/
var Excelify = function(startupOptions){
  this.containerElm = null; //reqd
  this.rowSelector = '.row';
  this.cellSelector = '.excelcell'; // should only select cells with input fields inside, not headings
  this.cellInputClassName = 'cellinput'; // should be unique such that cellSelector does not contain this
  this.cellInputQuerySelector = 'AUTO'; // "AUTO" to enable auto compute from '.'+cellInputClassName
  this.selectionBorderStyle = '2px solid #1567F0';
  this.copiedSelectionBorderStyle = '2px dashed #1567F0';
  this.selectionBackgroundStyle = '#C5E0FF';
  this.copyAreaSelector = null; //to have div on page defined that contains <textarea> for copy and paste intercept
  this.headingClassName = '';
  this.headingQuerySelector = 'AUTO';
  this.buttonBar = null;
  this.skipInvisibleCells = true;
  this.singleCellEditingMode = false;

  this.resetState = function(){
    this.tableCellContainers = [];
    this.tableCells = [];
    this.activeCell = null;
    this.totalDimensions = {x: 0, y: 0};
    this.activeCellIndex = {x: 0, y: 0};
    this.isDragging = false;
    this.dragOrigin = {x: 0, y: 0}; // x = col, y = row
    this.selectionStart = {x: 0, y: 0};
    this.selectionEnd = {x: 0, y: 0};
    this.copySelectionStart = {x: 0, y: 0};
    this.copySelectionEnd = {x: 0, y: 0};
    this.curSelectionisCopySel = false;
  };

  this.resetState();

  var c,cl, r,rl, cell, evcell, x,y,xl,yl; // private counter vars

  this.init = function(startupOptions){
    this.autoProps(startupOptions);
    this.rebuildIndex();
    this.attachListeners();
    this.validate();
    this.historyUtils.applyStateFn = this.applyHistoryState.bind(this);
    this.clipboardUtils.copyAreaSelector = this.copyAreaSelector;
    this.setupButtonBar();
    this.storeStateInHistory();
  };

  this.autoProps = function(){
    for( var k in startupOptions ){ this[k] = startupOptions[k]; }
    var auto = ['cellInputQuerySelector', 'headingQuerySelector'];
    for( c=0, x=0, cl=auto.length; c<cl; c++ ){
      var fromField = auto[c].replace(/QuerySelector$/, 'ClassName');
      if( this[auto[c]] == "AUTO" && this[fromField] ){
        this[auto[c]] = '.'+this[fromField];
      }else{
        this[auto[c]] = null;
      }
    }
  };

  this.validate = function(){
    if( this.cellInputClassName.indexOf('.') > -1 ){
      console.error('Excelify:Validation:Failure : cellInputClassName '+this.cellInputClassName+' looks like a selector, should be a unique class name');
    }
    if( this.headingClassName.indexOf('.') > -1 ){
      console.error('Excelify:Validation:Failure : headingClassName '+this.headingClassName+' looks like a selector, should be a unique class name');
    }
    if( this.cellSelector.indexOf(this.cellInputClassName) > -1 ){
      console.error('Excelify:Validation:Failure : cellInputClassName '+this.cellInputClassName+' must not be found in cellSelector '+this.cellSelector);
    }
  };

  // call this function whenever the table dom has been modified in a way that there is now a new spreadsheet or new visibility of cells/rows, update all cell references.
  this.rebuildIndex = function(){
    //before rebuilding index we can try to clear previous selection, but only if there was a previous index
    if( this.tableCells[0] && this.tableCells[0].length ){
      this.hideCurrentSelection();
      this.hideCopySelection();
    }
    this.resetState();
    // without rows we can still build the index from cells using an offset position helper function
    var rows = this.containerElm.querySelectorAll(this.rowSelector), cells, cellContainer, colHeadings=[], rowHeadings=[], rheading, cheadings;
    if( this.headingQuerySelector ){
      cheadings = rows[0].querySelectorAll(this.headingQuerySelector);
      for( c=0, x=0, cl=cheadings.length; c<cl; c++ ){
        if( this.skipInvisibleCells && !this.elementIsVisible(cheadings[c]) ) continue;
        cheadings[c].setAttribute('data-excelify-col', x++);
        cheadings[c].setAttribute('data-excelify-row', '*');
        colHeadings.push(cheadings[c]);
      }
    }

    for( r=0, y=0, rl=rows.length; r<rl; r++ ){
      this.tableCellContainers[y] = [], this.tableCells[y] = [];
      if( this.skipInvisibleCells && !this.elementIsVisible(rows[r]) ) continue;
      cells = rows[r].querySelectorAll(this.cellSelector);
      for( c=0, x=0, cl=cells.length; c<cl; c++ ){
        cellContainer = cells[c];
        if( this.skipInvisibleCells && !this.elementIsVisible(cellContainer) ) continue;
        this.tableCellContainers[y][x] = cellContainer; // store cell container
        cell = cells[c].querySelector(this.cellInputQuerySelector);
        if( cell ){
          this.tableCells[y][x] = cell;
          cell.setAttribute('data-excelify-col', x++);
          cell.setAttribute('data-excelify-row', y);
        }
      }
      if( this.tableCellContainers[y].length ){
        if( this.headingQuerySelector ){
          rheading = rows[r].querySelector(this.headingQuerySelector);
          rheading.setAttribute('data-excelify-col', '*');
          rheading.setAttribute('data-excelify-row', y);
          if( rheading ) rowHeadings.push(rheading);
        }
        y++;
      }
    }
    this.totalDimensions = {x: this.tableCells[0] ? this.tableCells[0].length-1 : 0, y: y-1};
  };

  this.setupButtonBar = function(){
    if( this.buttonBar ){
      this.historyUtils.buttonBarElements = {
        undo: this.buttonBar.querySelector('.undo'),
        redo: this.buttonBar.querySelector('.redo')
      };
      this.buttonBar.querySelector('.undo').addEventListener('click', this.historyUtils.undo.bind(this.historyUtils));
      this.buttonBar.querySelector('.redo').addEventListener('click', this.historyUtils.redo.bind(this.historyUtils));
    }
  };

  this.attachListeners = function(){ // only call again if container element changes
    this.containerElm.addEventListener('mousedown', this.mouseDownContainer.bind(this));
    this.containerElm.addEventListener('mouseup', this.mouseUpContainer.bind(this));
    this.containerElm.addEventListener('mouseover', this.mouseMoveContainer.bind(this));
    // key events only fire when a text cell is currently active
    document.addEventListener('keydown', this.keyboardDnEvents.bind(this));
    document.addEventListener('keyup', this.keyboardUpEvents.bind(this));
  };

  this.keyboardDnEvents = function(ev){
    if( ev.metaKey ){ // command/control
      switch(ev.keyCode){
        case 67: // C key - Copy
          this.captureCellCopy(ev);
          return;
        case 86: // V key - Paste
          this.applyCellPaste();
          return;
        case 88: // X key - Cut
          this.captureCellCopy(ev);
          this.setValueMultiCell(this.selectionStart, this.selectionEnd, ''); 
          return;
        case 90: // Z key - Undo / Redo
          if( ev.shiftKey ){
            this.historyUtils.redo(); // Cmd-Shift-Z redo
          }else{
            this.historyUtils.undo(); // Cmd-Z undo
          }
          return;
      }
      if( ev.charCode-0 === 0 ){
        this.prepareClipboardOverlay();
      }
    }else{
      switch(ev.keyCode){
        case 27: // ESC key
          this.singleCellEditingMode = false;
          this.clipboardUtils.hideArea();
          return;
        case 46: // Delete key - Clear Cells
          this.setValueMultiCell(this.selectionStart, this.selectionEnd, ''); 
          return;
        case 9: // Tab key already moves cells to right, but now we save state on each key press
          this.storeStateInHistory();
          return;
        case 13: // Enter key - move to next cell
          this.moveToNextCell();
          return;
      }
    }
  };

  this.keyboardUpEvents = function(ev){
    if( this.clipboardUtils.hideArea() ){
      setTimeout(this.activatePreviousCell.bind(this), 10);
    }
  };

  // you may wish override this functionality so the return key does something else!
  this.moveToNextCell = function(){
    var selSize = this.selectionSize();
    if( selSize.total > 1  ){
      this.activeCellIndex.y++; 

      if( this.activeCellIndex.y > this.selectionEnd.y ){
        this.activeCellIndex.y = this.selectionStart.y;
        this.activeCellIndex.x += 1;
        if( this.activeCellIndex.x > this.selectionEnd.x ){
          this.activeCellIndex.x = this.selectionStart.x;
        }
      }
      this.activeCell = this.tableCells[this.activeCellIndex.y][this.activeCellIndex.x];
      this.activeCell.select();

    }else{
      // if selections size is zero, move down one cell
      if( this.tableCells[this.activeCellIndex.y+1] ){
        this.activeCell = this.tableCells[this.activeCellIndex.y+1][this.activeCellIndex.x];
        this.activeCellIndex.y += 1;
        this.activeCell.select();
      }
    }
    if( !this.isDragging ) this.storeStateInHistory(); // in case we made a change and pressed return
  };

  this.elementIsVisible = function(cell){
    return  cell.clientWidth !== 0 && cell.clientHeight !== 0 && //cell.style.opacity !== 0 &&
            cell.style.visibility !== 'hidden';
  };

  this.cellPosition = function(cell){
    return {
      x: cell.getAttribute('data-excelify-col') - 0,
      y: cell.getAttribute('data-excelify-row') - 0
    };
  };

  this.mouseDownContainer = function(ev){
    var evcell = ev.target;
    if( evcell.className.indexOf(this.cellInputClassName) < 0 ){
      if( evcell.className.indexOf(this.headingClassName) < 0 ){
          this.hideCurrentSelection();
          return;
      }
    }
    if( document.releaseCapture ){
      setTimeout(function(){ // Firefox support
        document.releaseCapture();
      }, 10);
    }
    this.isDragging = true;
    this.activeCell = evcell;
    this.dragOrigin = this.cellPosition(evcell);
    if( this.pointsEqual(this.activeCellIndex, this.dragOrigin) ){
      this.singleCellEditingMode = true;
    }
    this.activeCellIndex = this.cellPosition(evcell);
    this.mouseMoveContainer(ev);
  };

  this.mouseUpContainer = function(ev){
    this.isDragging = false;
    var evcell = ev.target;
    if( evcell.className.indexOf(this.cellInputClassName) < 0 ){
      return;
    }
    this.storeStateInHistory(); // in case we just made a change
    var endPosition = this.cellPosition(ev.target);
    if( !this.singleCellEditingMode ){
      var selSize = this.subtractPoints(endPosition, this.activeCellIndex);
      if( selSize.total == 1 ){
        var cursorSelSize = this.activeSelectionSize();
        if( !cursorSelSize ){
          ev.target.select();
        }else{
          this.singleCellEditingMode = true;
        }
      }
    }
  };

  this.mouseMoveContainer = function(ev){
    if( this.isDragging ){
      if( ev.which === 0 ){
        this.isDragging = false; // we cancel the drag if we are back over container but the mouse button is not down
        return;
      }
      var evcell = ev.target, evinput = ev.target, currentPosition;
      if( evcell.className.indexOf(this.cellInputClassName) < 0 ){
        evinput = evcell.querySelector(this.cellInputQuerySelector);
        if( !evinput ){
          if( evcell.className.indexOf(this.headingClassName) ){
            currentPosition = this.cellPosition(evcell);
            this.selectBoxedCells(this.dragOrigin, currentPosition);
          }
          return; // in case user is still dragging, do not cancel until the mouse returns
        }
      }
      currentPosition = this.cellPosition(evinput);
      if( this.singleCellEditingMode && !this.pointsEqual(currentPosition, this.activeCellIndex) ){
        this.singleCellEditingMode=false; // so much for single cell editing mode, the active cell has changed
      }else{
        this.boxCells(this.dragOrigin, currentPosition); // if single editing this is superfluous
      }
    }
  };

  this.applyHistoryState = function(stateData){
    this.applyingHistoryState = true;
    this.setMultiValueMultiCell({x:0, y:0}, stateData);
    this.applyingHistoryState = false;
  };

  this.storeStateInHistory = function(){
    if( !this.applyingHistoryState ){
      this.historyUtils.addState(this.getAllCellValues());
    }
  };

  // the idea here is to capture unique states, and support undo and redo states
  this.historyUtils = {
    historyStates:  [],
    maxStates: 50,
    stateIndex: -1,
    buttonBarElements: null,
    applyStateFn: function(){},
    addState: function(data){
      var index = this.historyStates.length;
      if( this.stateIndex < index-1 ){
        this.historyStates.splice(this.stateIndex+1);
      }
      data = JSON.stringify(data);
      if( this.historyStates[this.stateIndex] != data ){// if data changed, store it!
        this.historyStates.push(data);
        if( this.historyStates.length > this.maxStates ){
          this.historyStates.splice(0, this.historyStates.length - this.maxStates); // cull old states
        }
        this.stateIndex = this.historyStates.length - 1;
        this.buttonBarUpdate();
      }
    },
    undo: function(){
      this.stateIndex--;
      if( this.stateIndex < 0 ) this.stateIndex = 0;
      this.applyStateFn(JSON.parse(this.historyStates[this.stateIndex]));
      this.buttonBarUpdate();
    },
    redo: function(){
      this.stateIndex++;
      if( this.stateIndex >=this.historyStates.length -1 ) this.stateIndex = this.historyStates.length -1;
      this.applyStateFn(JSON.parse(this.historyStates[this.stateIndex]));
      this.buttonBarUpdate();
    },
    buttonBarUpdate: function(){
      if( this.buttonBarElements ){
        var bb = this.buttonBarElements;
        if( this.stateIndex >= this.historyStates.length -1 ){
          bb.redo.style.opacity = '0.5';
        }else{
          bb.redo.style.opacity = '1.0';
        }
        if( this.stateIndex <= 0 ){
          bb.undo.style.opacity = '0.5';
        }else{
          bb.undo.style.opacity = '1.0';
        }
      }
    }
  };

  //could configure this to use some excel like editing area(top of screen) instead of creating new text-area that flashes over
  this.clipboardUtils = {
    previouslyFocusedElement: null,
    textareaStyle: 'position:fixed;top:25%;left:25%;right:25%;width:50%;opacity:0.5',
    lastArea: null,
    isShowing: false,
    hideArea: function(){
      var wasShowing = this.isShowing;
      if( this.isShowing ){
        if( this.copyAreaSelector ){
          document.querySelector(this.copyAreaSelector).style.display="none";
        }else if(this.lastArea){
          this.lastArea.style.display="none";
        }
        this.isShowing=false;
      }
      return wasShowing;
    },
    showArea: function(cvalue){
      var n;
      if( this.copyAreaSelector ){
        n = document.querySelector(this.copyAreaSelector);
        n.style.display="block";
        n = n.querySelector('textarea');
      }else if( this.lastArea ){
        n = this.lastArea;
        n.style.display="block";
      }else{
        n=document.createElement('textarea');
        n.setAttribute('style',this.textareaStyle);
        document.body.appendChild(n);
        this.lastArea = n;
      }
      this.isShowing=true;
      if( cvalue ){
        n.value=cvalue;
        setTimeout(function(){n.select();}, 15);
      }else{
        n.value='';
      }
      return n;
    },
    getPaste: function getPaste(cbf){
      var n=this.showArea();
      n.focus();
      n.select();
      setTimeout(function(){
        cbf(n.value);
        this.hideArea();
      }.bind(this), 250); // excessive wait time for paste completion?
    }
  };

  this.getCurrentSelectionForCopy = function(){
      var start = this.selectionStart,
          end   = this.selectionEnd,
          clipb = '';
      for( y=start.y, yl=end.y+1; y<yl; y++ ){
        for( x=start.x, xl=end.x; x<xl; x++ ){
          clipb += this.tableCells[y][x].value+"\t";
        }
        clipb += this.tableCells[y][x].value+"\n"; // last element in row gets \n instead of \t
      }
      return clipb;
  };

  this.activeSelectionSize = function(){
    if( this.activeCell ) return this.activeCell.selectionEnd - this.activeCell.selectionStart;
    return 0;
  };

  this.captureCellCopy = function(ev){
    var cursorSelSize = this.activeSelectionSize();
    if( cursorSelSize < 1 ) this.singleCellEditingMode = false;
    if( this.singleCellEditingMode ) return; // not sure about this yet
    var clipb = this.getCurrentSelectionForCopy();
    this.hideCopySelection();
    this.copySelectionStart = {x: this.selectionStart.x, y: this.selectionStart.y};
    this.copySelectionEnd = {x: this.selectionEnd.x, y: this.selectionEnd.y};
    this.styleEdges(this.copySelectionStart, this.copySelectionEnd, this.copiedSelectionBorderStyle);
    this.styleCells(this.copySelectionStart, this.copySelectionEnd, '');
    this.curSelectionisCopySel = true;
  };

  this.prepareClipboardOverlay = function(ev){
    var selSize = this.selectionSize();
    var cursorSelSize = this.activeSelectionSize();
    if( selSize.total != 1 || !cursorSelSize ){
      // if we have more than one cell selected or if the current selection within the cell is empty, show copy area
      this.clipboardUtils.showArea(this.getCurrentSelectionForCopy());
    }
  };

  this.activatePreviousCell = function(){
    if(this.activeCell) this.activeCell.focus();
  };

  this.applyCellPaste = function(){
    this.clipboardUtils.getPaste(this.valuesPasted.bind(this));
  };

  this.assembleIndexedPaste = function(activeCell, v){ // designed to be over-ridden
    var val = activeCell.value;
    var newValue = val.slice(0, activeCell.selectionStart) + v + val.slice(activeCell.selectionStart, val.length);
    // here is where we might need to set the caret or cursor position back to what it was previously.
    // it should go to the previous position plus v.length
    activeCell.value = newValue;
  };

  this.valuesPasted = function(v){
    var pasted = [];
    var rows = v.split("\n"); // it should end with one \n followed by nothing
    var rowCount = 0;
    for( r=0, x=1, rl=rows.length; r<rl; r++,x++ ){
      if( rows[r].length < 1 && x == rl ) continue; // this was to capture last row...
      pasted[r] = [];
      cells = rows[r].split("\t");
      for( c=0, cl=cells.length; c<cl; c++ ){
        pasted[r][c] = cells[c];
      }
      rowCount++;
    }
    if( pasted.length == 1 && pasted[0].length == 1 ){ // determine size of paste is greater than one cell or not, if not perform default paste action
      if( this.singleCellEditingMode ){
        this.assembleIndexedPaste(this.activeCell, v);
        return;
      }
    }

    var selSize = this.selectionSize();
    if( selSize.total > 1 && (rowCount != selSize.y || cl != selSize.x) ){
      if( !this.selectionConfirmation(selSize, {x: cl, y: rowCount}) ){
        setTimeout(this.activatePreviousCell.bind(this), 250);
        return;
      }
    }
    if( pasted[0] ){
      this.hideCurrentSelection();
      this.selectionEnd = this.validateSelectionCoordinate({x: this.selectionStart.x + pasted[0].length-1, y: this.selectionStart.y + pasted.length-1});
      this.styleActiveSelection();
      this.styleEdges(this.copySelectionStart, this.copySelectionEnd, ''); // hide copy region after paste
    }
    this.setMultiValueMultiCell(this.selectionStart, pasted);

    setTimeout(this.activatePreviousCell.bind(this), 250);
  };

  this.selectionConfirmation = function(selSize, clipSize){ // override
    return confirm('Selection size ('+selSize.x+','+selSize.y+') mismatches clipboard size ('+clipSize.x+','+clipSize.y+'), continue paste?');
  };

  this.getAllCellValues = function(){
    var allValues = [];
    for( y=0,yl=this.tableCells.length; y<yl; y++ ){
      allValues[y] = [];
      for( x=0,xl=this.tableCells[y].length; x<xl; x++ ){
        allValues[y][x] = this.tableCells[y][x].value;
      }
    }
    return allValues;
  };

  this.setMultiValueMultiCell = function(start, values){
    for( y=0,yl=values.length; y<yl; y++ ){
      for( x=0,xl=values[y].length; x<xl; x++ ){
        cell = this.tableCells[y+start.y];
        if( cell ){
          cell = cell[x+start.x];
          if( cell ) cell.value = values[y][x];
        }
      }
    }
    this.storeStateInHistory();
  };

  this.setValueMultiCell = function(start, end, value){
    for( y=start.y,yl=end.y+1; y<yl; y++ ){
      for( x=start.x,xl=end.x+1; x<xl; x++ ){
        this.tableCells[y][x].value = value;
      }
    }
    this.storeStateInHistory();
  };

  this.styleCells = function(start, end, backgroundStyle){
    for( y=start.y,yl=end.y+1; y<yl; y++ ){
      for( x=start.x,xl=end.x+1; x<xl; x++ ){
        this.tableCells[y][x].style.background = backgroundStyle;
      }
    }
  };

  this.styleEdges = function(start, end, borderStyle){
    for( x=start.x, xl=end.x, y=start.y, yl=end.y; y<=yl; y++ ){
      this.drawBorder(this.tableCellContainers[y][x], 'left', borderStyle);
      this.drawBorder(this.tableCellContainers[y][xl], 'right', borderStyle);
    }
   for( y=start.y; x<=xl; x++ ){
      this.drawBorder(this.tableCellContainers[y][x], 'top', borderStyle);
      this.drawBorder(this.tableCellContainers[yl][x], 'bottom', borderStyle);
    }
  };

  this.drawBorder = function(cell, side, borderStyle){
    cell.style['border-'+side] = borderStyle;
  };

  this.validateStartCoord = function(c){
    if( isNaN(c.x) ) c.x = 0;
    if( isNaN(c.y) ) c.y = 0;
    return this.validateSelectionCoordinate(c);
  };

  this.validateEndCoord = function(c){
    if( isNaN(c.x) ) c.x = this.tableCells[0].length-1;
    if( isNaN(c.y) ) c.y = this.tableCells.length-1;
    return this.validateSelectionCoordinate(c);
  };

  this.validateSelectionCoordinate = function(coordinate){
    if( coordinate.y > this.tableCells.length-1){
      coordinate.y = this.tableCells.length-1;
    }
    if( coordinate.x > this.tableCells[0].length-1){
      coordinate.x = this.tableCells[0].length-1;
    }
    if( coordinate.y < 0 ) coordinate.y = 0;
    if( coordinate.x < 0 ) coordinate.x = 0;
    return coordinate;
  };

  this.styleActiveSelection = function(){
      this.styleCells(this.selectionStart, this.selectionEnd, this.selectionBackgroundStyle);
      this.styleEdges(this.selectionStart, this.selectionEnd, this.selectionBorderStyle);
  };

  this.hideCurrentSelection = function(){
    this.styleEdges(this.selectionStart, this.selectionEnd, '');
    this.styleCells(this.selectionStart, this.selectionEnd, '');
  };

  this.hideCopySelection = function(){
    this.styleEdges(this.copySelectionStart, this.copySelectionEnd, '');
    this.styleCells(this.copySelectionStart, this.copySelectionEnd, '');
  };

  this.selectColumn = function(index){
    this.selectBoxedCells(
      {x: index, y: 0},
      {x: index, y: this.totalDimensions.y}
    );
  };

  this.selectRow = function(index){
    this.selectBoxedCells(
      {x: 0, y: index},
      {x: this.totalDimensions.x, y: index}
    );
  };

  this.selectAll = function(){
    this.selectBoxedCells(
      {x: 0, y: 0},
      {x: this.totalDimensions.x, y: this.totalDimensions.y}
    );
  };

  this.selectBoxedCells = function(startPos, endPos){
    startPos = this.validateStartCoord(startPos);
    endPos = this.validateEndCoord(endPos);
    this.activeCell = this.tableCells[startPos.y][startPos.x];
    if( this.activeCell ){
      this.activeCellIndex = this.cellPosition(this.activeCell);
      this.boxCells(startPos, endPos);
      var activeCell = this.activeCell;
      setTimeout(function(){activeCell.focus();}, 10);
    }
  };

  this.boxCells = function(startPos, endPos){ // also orders start/end position for us
    var startx = startPos.x, endx = endPos.x, starty = startPos.y, endy = endPos.y;
    if( startPos.x > endPos.x ){
      startx = endPos.x, endx = startPos.x;
    }
    if( startPos.y > endPos.y ){
      starty = endPos.y, endy = startPos.y;
    }
    if( this.curSelectionisCopySel ){
      this.curSelectionisCopySel = false;
    }else{
      this.hideCurrentSelection();
    }
    this.selectionStart = {x: startx, y: starty};
    this.selectionEnd = {x: endx, y: endy};
    this.styleActiveSelection();
  };

  this.subtractPoints = function(end, start){
    var obj = {
      x: end.x+1 - start.x,
      y: end.y+1 - start.y
    };
    obj.total = Math.abs(obj.x * obj.y);
    return obj;
  };

  this.pointsEqual = function(a, b){
    return a.x == b.x && a.y == b.y;
  };

  this.selectionSize = function(){
    return this.subtractPoints(this.selectionEnd, this.selectionStart);
  };

  this.init(startupOptions); // new Excelify(startupOptions);
};
