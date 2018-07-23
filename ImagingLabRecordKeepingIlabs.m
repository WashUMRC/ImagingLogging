function ImagingLabRecordKeepingIlabs()

outFile = ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\User Logs\UserLogIlabs.txt'];
numLogs = length(dir(['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\User Logs\Backups\*.txt']));

prompt={'Your wustl email address','PI wustl email address','Is your PI a core member? (y/n)','Sample Number (if not CT scanning, enter 0)','Is this a live animal test? (y/n)'};
dlgTitle='User Inputs';
lineNo=1;
answer=inputdlg(prompt,dlgTitle,lineNo);
for i = 1:length(answer)
    if isempty(answer{i})
        errordlg('Please do not leave any fields blank!');
        return
    end
end

global task;

taskGUI();
task = 0;
while task == 0
    pause(0.1);
end


user = answer{1};
PI = answer{2};
member = answer{3};
sample = answer{4};
live = answer{5};
clear answer

if task == 1
    theTask = 'mictoCT 40 scanning';
elseif task == 2
    theTask = 'vivaCT scanning';
elseif task == 3
    theTask = 'Dan Time';
    answer = inputdlg('Please enter the amount of technician time used in hours:');
    if isempty(answer{1})
        errordlg('Please enter a quantity!')
        return
    end
    answer2 = inputdlg('Please enter a description of what you were doing');
elseif task == 4
    theTask = 'Michael Time';
    answer = inputdlg('Please enter the amount of technician time used in hours:');
    if isempty(answer{1})
        errordlg('Please enter a quantity!')
        return
    end
    answer2 = inputdlg('Please enter a description of what you were doing');
end

if exist('answer') == 1
    types = whos('answer');
    if isempty(strfind(types.class,'double'))
        quantity = str2num(answer{1});
    else
        quantity = answer;
    end
end
if exist('answer2') == 1
    description = answer2{1};
end

fid = fopen(outFile,'a');
fprintf(fid,'%s\t',datestr(now));
fprintf(fid,'%s\t',user);
fprintf(fid,'%s\t',PI);
fprintf(fid,'%s\t',member);
fprintf(fid,'%s\t',sample);
fprintf(fid,'%s\t',live);
fprintf(fid,'%s\t',theTask);
if exist('quantity') == 1
    fprintf(fid,'%s\t',num2str(quantity));
else
    fprintf(fid,'%s\t','');
end
if exist('description') == 1
    fprintf(fid,'%s\t',description);
end
fprintf(fid,'\n');


fclose(fid);
fclose('all');

sysLine = ['copy "J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\User Logs\UserLogIlabs.txt" ' '"J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\User Logs\backups\UserLogIlabs' num2str(numLogs+1) '.txt"'];
system(sysLine);

msgbox('Thank you! Our records have been updated.');

function taskGUI()

sz = [500 800]; % figure size
screensize = get(0,'ScreenSize');
xpos = ceil((screensize(3)-sz(2))/2); % center the figure on the screen horizontally
ypos = ceil((screensize(4)-sz(1))/2); % center the figure on the screen vertically
h.fig = figure('position',[xpos ypos sz(2) sz(1)],'Name','Select your task','NumberTitle','off');

h.h1 = uicontrol('Style','pushbutton','units','normalized','Position',[0.1 0.1 .33 .25],...
    'string','MicroCT','fontsize',14,'Callback','call1');

set(h.h1,'callback',{@call1, h});

h.h2 = uicontrol('Style','pushbutton','units','normalized','Position',[.6 .1 .33 .25],...
    'string','VivaCT','fontsize',14,'Callback','call2');

set(h.h2,'callback',{@call2, h});

h.h3 = uicontrol('Style','pushbutton','units','normalized','Position',[.1 .6 .33 .25],...
    'string','Dan','fontsize',14,'Callback','call3');

set(h.h3,'callback',{@call3, h});

h.h4 = uicontrol('Style','pushbutton','units','normalized','Position',[.6 .6 .33 .25],...
    'string','Michael','fontsize',14,'Callback','call4');

set(h.h4,'callback',{@call4, h});

function h = call1(hObject,eventdata,h)
global task
task = 1;
close('all');

function h = call2(hObject,eventdata,h)
global task
task = 2;
close('all');

function h = call3(hObject,eventdata,h)
global task
task = 3;
close('all');

function h = call4(hObject,eventdata,h)
global task
task = 4;
close('all');
