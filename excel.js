/*
In comments, I have called out a handful of places where generic utility
code was gotten from online, as I would consider normal & acceptable.
I considered the column-letter-to-number functions generic problems not
crucial to solve on my own. I also found some boilerplate
for building a 'clickable' HTML grid in JS to give me a starting point for
writing the logic. All other code is my own.
*/


/* MY PLANNING & APPROACH
I recognised that the scope of the exercise would likely take me longer than
the asked-for 5-6 hours, but I also understood that spending significantly longer
is not the idea of the exercise. So, I aggressively limited the scope to getting
working versions of exactly what was asked for.

For steps 3, 5 and 6, formulas and functions, entry of strings containing non-numeric
characters was allowed so that column names could be entered without adding
the complexity of error handling. Handling of invalid inputs (including cells
referencing themselves or cells recursively referencing each other)
was judged out-of-scope.

The exact types of formulas and functions given as examples were prioritized as
the ones to get done, demonstrating a working starting point that further math
operations would build from. Handling of non-reference equations like '=1+2'
was disregarded as not-to-spec.

I have given the code a readability pass but have left alone
low-priority refactoring opportunities such as making the formatting
functionality in step 7 look "nicer" in code.
In total I spent about 7 hours on this, plus time for adding these explanatory
comments and TODOs.
*/


// Globals
var lastClicked;
var rowCount = 100;
var columnCount = 100;
// TODO: refactor header row and column to not affect the
// rest of the grid and not cause this off-by-one problem
var formulaList = new Array(rowCount * columnCount + 1).fill(null);
var evaluatedList = new Array(rowCount * columnCount + 1).fill(null);


function loadClickableGrid(
    rows, cols, evaluatedList, onClickCallback, focusOutCallback
) { // https://jsfiddle.net/6qkdP/2/
    /* Load or reload grid and cells */
    var grid = document.createElement('table');
    grid.className = 'grid';
    grid.id = 'grid';
    var header = grid.appendChild(document.createElement('tr'));
    // Insert the empty cell at the top left
    var empty = header.appendChild(document.createElement('td'));
    var cellNum = 1;

    for (var columnNum = 0; columnNum < cols; columnNum++) {
        var columnLabel = header.appendChild(document.createElement('td'));
        // TODO: off-by-one
        columnLabel.innerHTML = colName(columnNum);
        columnLabel.style.fontWeight = 'bold';
    };

    for (var rowNum = 1; rowNum <= rows; rowNum++) {
        var row = grid.appendChild(document.createElement('tr'));
        var rowLabel = row.appendChild(document.createElement('td'));
        rowLabel.innerHTML = rowNum;
        rowLabel.style.fontWeight = 'bold';

        for (var columnNum = 1; columnNum <= cols; columnNum++) {
            var cell = row.appendChild(document.createElement('td'));
            cell.contentEditable = 'true';
            cell.id = 'cell' + cellNum.toString();
            cell.innerHTML = evaluatedList[cellNum];

            cell.addEventListener('click', (function (cell, rowNum, colNum, cellNum) {
                return function () { onClickCallback(cell, rowNum, colNum, cellNum); }
            })(cell, rowNum, columnNum, cellNum), false);

            cell.addEventListener('focusout', (function (cell, cellNum) {
                return function () { focusOutCallback(cell, cellNum); }
            })(cell, cellNum), false);

            cellNum++;
        }
    }
    return grid;
};

// https://stackoverflow.com/questions/8240637/convert-numbers-to-letters-beyond-the-26-character-alphabet
function colName(n) {
    // Get column name from number, e.g. 0 = A, 26 = AA
    var ordA = 'A'.charCodeAt(0);
    var ordZ = 'Z'.charCodeAt(0);
    var len = ordZ - ordA + 1;

    var s = "";
    while (n >= 0) {
        s = String.fromCharCode(n % len + ordA) + s;
        n = Math.floor(n / len) - 1;
    }
    return s;
};

// https://stackoverflow.com/questions/9905533/convert-excel-column-alphabet-e-g-aa-to-number-e-g-25
function lettersToNumber(letters) {
    // Get number from column name, e.g. A = 1, AA = 27
    for (var p = 0, n = 0; p < letters.length; p++) {
        n = letters[p].charCodeAt() - 64 + n * 26;
    }
    return n;
};

// https://jsfiddle.net/6qkdP/2/
function cellClicked(cell, row, col, cellNum) {
    console.log("You clicked on cell:", cell);
    console.log("You clicked on row:", row);
    // TODO: off-by-one
    console.log("You clicked on col:", colName(col - 1));
    console.log("You clicked on cell number #:", cellNum);

    cell.className = 'clicked';
    if (lastClicked) lastClicked.className = '';
    lastClicked = cell;
    // Replace evaluated contents with original formula when clicked
    cell.innerHTML = formulaList[cellNum];
};

function onFocusOut(cell, cellNum) {
    /* When clicking away from a cell, reevaluate all cells for
    possible effects on formulas */

    formulaList[cellNum] = cell.innerHTML;
    evaluateCells();
};

function evaluateCells() {
    /* Evaluate cell contents and update the list with their
    contents or output of their formulas. */
    for (i = 1; i < formulaList.length; i++) {
        evaluateCell(i);
    };
    // This accounts for formulas placed above the cells they reference.
    // TODO: this could cover everything only once by fanning out in both directions
    // from an edited cell, but this was the last thing I added before stopping.
    for (i = formulaList.length - 1; i > 0; i--) {
        evaluateCell(i);
    };
};

function evaluateCell(i) {
    if (formulaList[i]) {
        cell = document.getElementById('cell' + i.toString());
        cellContent = evaluateCellContent(formulaList[i]);
        cell.innerHTML = cellContent;
        evaluatedList[i] = cellContent;
    }
};

function evaluateCellContent(cellContent) {
    /* Figure out if the cell contents are a formula and calculate it if needed */
    if (!cellContent) {
        return null;
    };
    if (cellContent.startsWith('=') && cellContent.length > 1) {
        var formula = cellContent.substring(1);
        try {
            var result = calculateCellValue(formula);
            return result;
        } catch (SyntaxError) {
            // an incomplete formula like '=A1+'
            return cellContent;
        }
    }
    else {
        return cellContent;
    };
};

function calculateCellValue(formula) {
    /* Parse the expression in the cell and calculate the result */
    var expressionParts = getExpressionParts(formula);

    // TODO: pull out into a function which figures out what kind of
    // math we're doing
    if (expressionParts.includes('+')) {
        return calculateCellFormula(expressionParts);
    };
    if (expressionParts.includes(':')) {
        return calculateCellFunctionSum(expressionParts);
    };
};

function getExpressionParts(formula) {
    /* Figure out what kind of function or formula is in the cell */

    // TODO: this is where explicit operators begin to be used to send the
    // cell contents down different code paths. This needs further
    // abstraction and separation of concerns if adding more
    // operators (e.g. '*', AVERAGE), to figure out which kind of
    // math we're doing.

    if (formula.includes('+')) {
        // split but keep the '+'
        return formula.split(/(\+)/);
    }
    // TODO
    // if (formula.includes('*')) {
    //     return formula.split(/(\*)/);
    // }

    // TODO: pull out into a function which figures out what kind of
    // excel function we're doing. This code path is currently hardcoded
    // to know it's doing a SUM.
    if (formula.includes('sum')) {
        var expression = formula.split('(')[1].split(')')[0]
        // TODO: pull out into a function which figures out what kind of
        // excel delimiter we're using.
        if (expression.includes(':')) {
            // split but keep the ':'
            return expression.split(/(\:)/);
            // TODO
            // } else if (formula.includes(',')) {
            //     return expression.split(/(\,)/);
            // }
        }
        else {
            throw new SyntaxError('Formula malformed');
        };
    }
};

function calculateCellFormula(expressionParts) {
    /* Calculate the result for the formula in the cell */
    var expression = new Array();

    for (var i = 0; i < expressionParts.length; i++) {
        part = expressionParts[i];
        // TODO: pull out into a function which figures out which kind of
        // math we're doing
        if (part.includes('+')) {
            expression.push(part);
        } else {
            var letters = part.replace(/[0-9]/g, '');
            var numbers = part.replace(/([a-zA-Z])/g, '');
            var cellNum = flattenGridCoordinates(
                lettersToNumber(letters), Number(numbers)
            );
            expression.push(evaluatedList[cellNum]);
        }
    }
    return eval(expression.join(''));
};

function calculateCellFunctionSum(expressionParts) {
    /* Calculate the result for the function in the cell
    /* At this point we know we're SUMMING cells in a range like A1:A5

    TODO: make this work along the x-axis? e.g. sum(A1:B1) */
    var expression = new Array();
    firstPart = expressionParts[0];
    secondPart = expressionParts[expressionParts.length - 1];

    // pull out the column we're working with
    var firstPartLetters = firstPart.replace(/[0-9]/g, '');
    // starting row
    var firstPartNumbers = firstPart.replace(/([a-zA-Z])/g, '');
    // last row
    var secondPartNumbers = secondPart.replace(/([a-zA-Z])/g, '');

    // loop down the rows in the given column,
    // collecting the cell values within the range
    for (var i = firstPartNumbers; i <= secondPartNumbers; i++) {
        var cellNum = flattenGridCoordinates(
            lettersToNumber(firstPartLetters), Number(i)
        );
        cellValue = evaluatedList[cellNum];
        if (cellValue) {
            expression.push(cellValue);
        }
    }
    return eval(expression.join('+'));
};

function flattenGridCoordinates(rowNum, colNum) {
    // this formula flattens 2D to 1D for cellNum identification
    return rowNum + ((colNum - 1) * columnCount);
};

function initializeFormatters() {
    // Initialize the toggleable formatters
    document.getElementById('bold').addEventListener('click', (function () {
        console.log("You clicked on bold");
        if (lastClicked && lastClicked.style.fontWeight != 'bold') {
            lastClicked.style.fontWeight = 'bold';
        } else if (lastClicked) lastClicked.style.fontWeight = null;
    }));
    document.getElementById('italic').addEventListener('click', (function () {
        if (lastClicked && lastClicked.style.fontStyle != 'italic') {
            lastClicked.style.fontStyle = 'italic';
        } else if (lastClicked) lastClicked.style.fontStyle = null;
    }));
    document.getElementById('underline').addEventListener('click', (function () {
        console.log("You clicked on underline");
        if (lastClicked && !lastClicked.style.borderBottom) {
            lastClicked.style.borderBottom = '2px solid #000000';
        } else if (lastClicked) lastClicked.style.borderBottom = null;
    }));
    // Initialize reset
    document.getElementById('reset').addEventListener('click', (function () {
        console.log("You clicked on reset");
        var grid = document.getElementById('grid');
        grid.remove();
        var grid = loadClickableGrid(
            rowCount, columnCount, evaluatedList, cellClicked, onFocusOut
        );
        document.body.appendChild(grid);
    }));
};

function initializePage(rowCount, columnCount, evaluatedList) {
    var grid = loadClickableGrid(
        rowCount, columnCount, evaluatedList, cellClicked, onFocusOut
    );
    document.body.appendChild(grid);
    initializeFormatters();
};

initializePage(rowCount, columnCount, evaluatedList, document);
