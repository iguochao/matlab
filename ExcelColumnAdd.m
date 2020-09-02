function colCharY = ExcelColumnAdd(colCharX, n)
% determine the alphabetic characters of column index at a relative position in Excel sheet
% 
% GuoChao@ntu.edu.cn
% 2-Sep-2020
% %
% Input : 
% 1. colCharX: original alphabetic characters
% 2. n: the number to be added
% 'A': 1; 'B': 2; ...; 'Z':26; 'AA': 27; ...
%
% % Examples : 
% ExcelColumnAdd('A', 26)   % ans = 'AA' 
%
% e.g. when you determine the range in a Excel sheet for a datatable dt
% if the upper left cell is 'B2'， then you get range4dt:
% [nrow, ncol] = size(dt)
% range4dt=sprintf('%s%d:%s%d', 'B', 2, charAdd('B', ncol - 1), 2 + nrow - 1);
% eSheetx = eSheet1.get('Range', range4dt);
% 

colCharY = chard(dcharx(colCharX)+n);
end

function dy = dcharx(colCharX)
% convert colChar to number 
% test 'AA' 27
dcolCharX = double(colCharX);
dbase = double('A')-1;
if numel(dcolCharX) == 1
    dy = dcolCharX - dbase;
else
    dy = 26*dcharx(char(dcolCharX(1:end-1))) + dcolCharX(end)- dbase;
end
end

function colCharY = chard(dx)
% convert number to colChar
% check 26 Z, 52 AZ, 78 BZ

colCharY = [];
dbase = double('A');
if dx <= 26
    colCharY = char(dbase+dx-1);
else
    rem = mod(dx, 26);
    dx = fix(dx/26);
    if rem==0
        colCharY = [chard(dx-1) 'Z'];
    else
    colCharY = [chard(dx) char(rem+dbase-1)];
    end
end
end

