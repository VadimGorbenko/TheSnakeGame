(function initGame(){
    var field = document.querySelector('.snake-playground');
    var fieldCTX = field.getContext('2d');
    var fieldW = field.width;
    var fieldH = field.height;
    var cellSize = 10;
    var fieldWidthInCells = fieldW / cellSize;
    var fieldHeightInCells = fieldH / cellSize;
    var score = 0;

    var drowBorder = function(){
        fieldCTX.fillStyle = 'Gray';
        fieldCTX.fillRect(0, 0, fieldW, cellSize);
        fieldCTX.fillRect(0, fieldH - cellSize, fieldW,      cellSize);
        fieldCTX.fillRect(0, 0, cellSize, fieldH);
        fieldCTX.fillRect(fieldW - cellSize, 0, cellSize, fieldH);
    };

    var drawScore = function(){
        fieldCTX.font = '20px Courier';
        fieldCTX.fillStyle = 'Black';
        fieldCTX.textAlign = 'left';
        fieldCTX.textBaseline = 'top';
        fieldCTX.fillText('Счёт: ' + score, cellSize, cellSize);
    }
})()