% function completeImagingLabBillingCompilerIlabs()
% 
clear all;
clc;
qcNumber = 451;
% 
answer = inputdlg('Please enter the three letter abbreviation for the month of interest (or ALL for complete report)');
month = answer{1};
firstDate = datenum(['1-' month '-18']);
answer = inputdlg('Please enter the three letter abbreviation for the month AFTER the month of interest (or ALL for complete report)');
endMonth = answer{1};
lastDate = datenum(['1-' endMonth '-18']);
% 
%do Viva stuff first
h = msgbox('Do not forget to copy the two Faxitron databases to the server! You can go do that now until you close the box.');
uiwait(h);

% vivaCTTimeCollector();%collects time from rsq log txt files

dxaDir = ['J:\Silva''s Lab\P30 Core Center\Faxitron Backup\DXA\Data'];
dxaFiles = dir([dxaDir '\*.dcm']);

%include info for Faxitron data
c=0;
for i = 1:length(dxaFiles)
    clc
    i/length(dxaFiles)
    info = dicominfo([dxaDir '\' dxaFiles(i).name]);
    c=c+1;
    out{c,1} = datestr(datenum([info.AcquisitionDate(5:6) '\' info.AcquisitionDate(7:8) '\' info.AcquisitionDate(1:4)]));
    out{c,2} = [info.OperatorName.FamilyName];%technician
    out{c,3} = [info.ReferringPhysicianName.FamilyName];%PI
    out{c,4} = info.Filename;%image number
    out{c,5} = '';
    out{c,6} = '';
    dist(c) = info.DistanceSourceToDetector;
    if round(info.DistanceSourceToDetector) == 318
        out{c,7} = 'DEXA';
        out{c,8} = '';
        out{c,9} = '';
        out{c,10} = 'DEXA';
        out{c,11} = 1;
        out{c,12} = '';
        out{c,13} = '';
        out{c,14} = '';
        out{c,15} = '';
        out{c,16} = '';
        out{c,17} = '';
        out{c,18} = '';
        out{c,19} = '';
        out{c,20} = '';
        out{c,21} = '';
        out{c,22} = '';
        out{c,23} = '';
        out{c,24} = '';
        out{c,25} = '';
        out{c,26} = [info.PatientName.GivenName];
    else
        out{c,7} = 'X-Ray';
        out{c,8} = '';
        out{c,9} = '';
        out{c,10} = 'X-Ray';
        out{c,11} = 1;
        out{c,12} = '';
        out{c,13} = '';
        out{c,14} = '';
        out{c,15} = '';
        out{c,16} = '';
        out{c,17} = '';
        out{c,18} = '';
        out{c,19} = '';
        out{c,20} = '';
        out{c,21} = '';
        out{c,22} = '';
        out{c,23} = '';
        out{c,24} = '';
        out{c,25} = '';
        out{c,26} = [info.PatientName.GivenName];
    end
end

radiographDir = ['J:\Silva''s Lab\P30 Core Center\Faxitron Backup\Bioptics\Data'];
radiographFiles = dir([radiographDir '\*.dcm']);

for i = 1:length(radiographFiles)
    clc;
    i/length(radiographFiles)
    info = dicominfo([radiographDir '\' radiographFiles(i).name]);
    c=c+1;
    out{c,1} = datestr(datenum([info.AcquisitionDate(5:6) '\' info.AcquisitionDate(7:8) '\' info.AcquisitionDate(1:4)]));
    out{c,2} = [info.OperatorName.FamilyName];%technician
    out{c,3} = [info.ReferringPhysicianName.FamilyName];%PI
    out{c,4} = info.Filename;%image number
    out{c,5} = '';
    out{c,6} = '';
    if round(info.DistanceSourceToDetector) ~= 318
        out{c,7} = 'X-Ray';
        out{c,8} = '';
        out{c,9} = '';
        out{c,10} = 'X-Ray';
        out{c,11} = 1;
        out{c,12} = '';
        out{c,13} = '';
        out{c,14} = '';
        out{c,15} = '';
        out{c,16} = '';
        out{c,17} = '';
        out{c,18} = '';
        out{c,19} = '';
        out{c,20} = '';
        out{c,21} = '';
        out{c,22} = '';
        out{c,23} = '';
        out{c,24} = '';
        out{c,25} = '';
        out{c,26} = [info.PatientName.GivenName];
    else
        out{c,7} = 'DEXA';
        out{c,8} = '';
        out{c,9} = '';
        out{c,10} = 'DEXA';
        out{c,11} = 1;
        out{c,12} = '';
        out{c,13} = '';
        out{c,14} = '';
        out{c,15} = '';
        out{c,16} = '';
        out{c,17} = '';
        out{c,18} = '';
        out{c,19} = '';
        out{c,20} = '';
        out{c,21} = '';
        out{c,22} = '';
        out{c,23} = '';
        out{c,24} = '';
        out{c,25} = '';
        out{c,26} = [info.PatientName.GivenName];
    end
end

%identify file of interest
foiOrtho = [['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Billing Records\'] month '18_Viva_Billing Information.tab'];
foiIlabs = [['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Billing Records\'] month '18_Viva_Billing Information iLabs.csv'];

userLogPath = ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\User Logs\UserLogIlabs.txt'];
vivaCompleteLogPath = ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\VivaCT Machine Logs\LogFromRSQHeader*.txt'];

excel = actxserver('Excel.Application');
set(excel,'Visible',0);
workbook = excel.Workbooks;
invoke(workbook,'Open',userLogPath);
excel.ActiveWorkbook.SaveAs([pwd '\userLogIlabs.xlsx'],51);
invoke(excel, 'Quit');
delete(excel);
[userNum,userTxt,userRaw] = xlsread([pwd '\userLogIlabs.xlsx']);
sysLine = (['del ' [pwd '\userLogIlabs.xlsx']]);
system(sysLine);

vivaFile = dir(vivaCompleteLogPath);
for i = 1:length(vivaFile)
    dates(i) = datenum(vivaFile(i).date);
end
[aa bb] = sort(dates,'descend');
vivaFile = vivaFile(bb);
vivaFile = vivaFile(1);
vivaRaw = importdata(['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\VivaCT Machine Logs\' vivaFile.name]);
[a b] = size(vivaRaw);
for i = 1:a
    for j = 1:b
        tmp{i,j} = vivaRaw(i,j);
    end
end
clear vivaRaw
vivaRaw = tmp;
clear tmp;

[a b] = size(userRaw);
for i = 2:a
    userRaw{i,3} = upper(userRaw{i,3});
end
[a b] = size(userRaw);
for i = 2:a
    userRaw{i,2} = upper(userRaw{i,2});
end
[d e] = size(vivaRaw);
for i = 2:a
    samples{i-1} = num2str(userRaw{i,5});
end
uniqueSamples = unique(samples);

%make a few corrections to userRaw
for i = 2:a
    if strcmpi(userRaw{i,4},'yes') == 1
        userRaw{i,4} = 'y';
    end
    if strcmpi(userRaw{i,4},'no') == 1
        userRaw{i,4} = 'n';
    end
    if strcmpi(userRaw{i,6},'yes') == 1
        userRaw{i,5} = 'y';
    end
    if strcmpi(userRaw{i,6},'no') == 1
        userRaw{i,5} = 'n';
    end
    userRaw{i,1} = datestr(datenum(userRaw{i,1}),1);
end

%Identify PI/Sample pairs
for i = 2:a
    PI{i-1} = userRaw{i,3};
    sample(i-1,:) = userRaw{i,5};
end
[uniqueSamples,uSampInd] = unique(sample);
for i = 1:length(uSampInd)
    PIs{i} = PI{uSampInd(i)};
end

userRaw = [userRaw;out(:,1:9)];

userRaw = userRawCorrection(userRaw);

%Unique Samples and PIs are now matched, arrange these knowns
c=0;
for i = 2:d
    for k = 2:a
        if ~ischar(vivaRaw{i,2}) &&  vivaRaw{i,2} == userRaw{k,5} && vivaRaw{i,2} < 1000
            c=c+1;
            matched{c,1} = datestr(datenum(vivaRaw{i,1}),1);%date of scan by machine log
            matched{c,2} = userRaw{k,3};%PI
            matched{c,3} = userRaw{k,4};%core member
            matched{c,4} = vivaRaw{i,2};%sample number
            matched{c,5} = vivaRaw{i,3};%measurement number
            matched{c,6} = userRaw{k,2};%user
            matched{c,7} = vivaRaw{i,4};%time
            matched{c,8} = userRaw{k,6};%live animal
        end
    end
end

c=0;
for i = 1:length(matched)
    if ~ischar(matched{i,5})
        c=c+1;
        measurements(c) = matched{i,5};
    end
end
uniqueMeasurements = unique(measurements);

%clean up duplication

cMicro=0;
for i = 1:length(uniqueMeasurements)
    cMicro=cMicro+1;
    for j = 1:length(matched)
        if matched{j,5} == uniqueMeasurements(i)
            matchedClean{cMicro,1} = matched{j,1};
            matchedClean{cMicro,2} = matched{j,2};
            matchedClean{cMicro,3} = matched{j,3};
            matchedClean{cMicro,4} = matched{j,4};
            matchedClean{cMicro,5} = matched{j,5};
            matchedClean{cMicro,6} = matched{j,6};
            matchedClean{cMicro,7} = matched{j,7};
            matchedClean{cMicro,8} = matched{j,8};
        end
    end
end

%Pull only measurements for month of interest
count=0;
for i = 1:cMicro
    if datenum(matchedClean{i,1}) >= firstDate && datenum(matchedClean{i,1}) < lastDate
        count=count+1;
        matchedInMonth{count,1} = matchedClean{i,1};
        matchedInMonth{count,2} = matchedClean{i,2};
        matchedInMonth{count,3} = matchedClean{i,3};
        matchedInMonth{count,4} = matchedClean{i,4};
        matchedInMonth{count,5} = matchedClean{i,5};
        matchedInMonth{count,6} = matchedClean{i,6};
        matchedInMonth{count,7} = matchedClean{i,7};
        matchedInMonth{count,8} = matchedClean{i,8};
    end
end


clear dates;
[a b] = size(matchedInMonth);
for i = 1:a
    dates{i} = matchedInMonth{i,1};
    PIs{i} = matchedInMonth{i,2};
end
uniqueDates = unique(dates);
uniquePIs = unique(PIs);

[a b] = size(matchedInMonth);
dayTotal=zeros(length(uniquePIs),length(uniqueDates));
c=0;
dayUser = cell(length(uniquePIs),length(uniqueDates));
for i = 1:length(uniqueDates)
    for j = 1:length(uniquePIs)
        for k = 1:a
            if strcmp(matchedInMonth{k,1},uniqueDates{i}) == 1 && strcmp(matchedInMonth{k,2},uniquePIs{j}) == 1 && matchedInMonth{k,4} ~= qcNumber
                dayTotal(j,i) = dayTotal(j,i) + matchedInMonth{k,7}; %end up with PI x day array of totals
                if isempty(dayUser{j,i})
                    dayUser{j,i} = matchedInMonth{k,6};
                elseif ~isempty(dayUser{j,i}) && strcmpi(dayUser{j,i},matchedInMonth{k,6}) ~= 1
                    dayUser{j,i} = [dayUser{j,i} '/' matchedInMonth{k,6}];
                end
            end
        end
    end
end

monthTotal = sum(sum(dayTotal));

sysLine = [['move /y "J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Billing Records\*Viva_Billing Information.tab'] '" "' ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Viva billing log backups\"']];
system(sysLine);

fidOrtho = fopen(foiOrtho, 'w');

% print headers
fprintf(fidOrtho,'%s\n',['VivaCT and other Core B billing report for ' month ' 2018']);
fprintf(fidOrtho,'%s\n','Reference list of all scan information');
fprintf(fidOrtho,'%s\t','Date');
fprintf(fidOrtho,'%s\t','PI');
fprintf(fidOrtho,'%s\t','Member');
fprintf(fidOrtho,'%s\t','Sample');
fprintf(fidOrtho,'%s\t','Measurement');
fprintf(fidOrtho,'%s\t','User');
fprintf(fidOrtho,'%s\t','Scan Time');
% fprintf(fidOrtho,'%s\t','Live');
fprintf(fidOrtho,'\n');

%print data on per-scan basis
[a b] = size(matchedInMonth);
for i = 1:a
    fprintf(fidOrtho,'%s\t',matchedInMonth{i,1});
    fprintf(fidOrtho,'%s\t',matchedInMonth{i,2});
    fprintf(fidOrtho,'%s\t',matchedInMonth{i,3});
    fprintf(fidOrtho,'%s\t',num2str(matchedInMonth{i,4}));
    fprintf(fidOrtho,'%s\t',num2str(matchedInMonth{i,5}));
    fprintf(fidOrtho,'%s\t',num2str(matchedInMonth{i,6}));
    fprintf(fidOrtho,'%s\t',matchedInMonth{i,7});
    if matchedInMonth{i,8} ~= 0
        hrs = floor(matchedInMonth{i,8});
        min = round(mod(matchedInMonth{i,8},1) * 60);
        fprintf(fidOrtho,'%s\t',[num2str(hrs) ':' num2str(min)]);
    elseif matchedInMonth{i,8} == 0
        fprintf(fidOrtho,'%s\t','Failed Scan');
    end
    %     fprintf(fidOrtho,'%s\n',matchedInMonth{i,9});
    fprintf(fidOrtho,'%s\n',' ');
end
fprintf(fidOrtho,'%s\n',' ');
fprintf(fidOrtho,'%s\n',' ');
fprintf(fidOrtho,'%s\n','Total Period Usage');
fprintf(fidOrtho,'%s\n',num2str(monthTotal));

fprintf(fidOrtho,'%s\n',' ');
fprintf(fidOrtho,'%s\n',' ');

%Print daily totals per PI
[a b] = size(dayTotal);
fprintf(fidOrtho,'%s\n','Daily totals by PI in hours:min');
for i = 1:a
    fprintf(fidOrtho,'%s\t',uniquePIs{i});
    for j = 1:b
        if dayTotal(i,j) ~= 0
            fprintf(fidOrtho,'%s\t',uniqueDates{j});
        end
    end
    fprintf(fidOrtho,'%s\n',' ');
    fprintf(fidOrtho,'%s\t',' ');
    for j = 1:b
        if dayTotal(i,j) ~= 0
            hrs = floor(dayTotal(i,j));
            min = round(mod(dayTotal(i,j),1) * 60);
            fprintf(fidOrtho,'%s\t',[num2str(hrs) ':' num2str(min)]);
        end
    end
    fprintf(fidOrtho,'%s\n',' ');
    fprintf(fidOrtho,'%s\t',' ');
    for j = 1:b
        if dayTotal(i,j) ~= 0
            fprintf(fidOrtho,'%s\t',dayUser{i,j});
        end
    end
    fprintf(fidOrtho,'%s\n',' ');
    dailyTotal = sum(dayTotal(i,:));
    hrs = floor(dailyTotal);
    min = round(mod(dailyTotal,1) * 60);
    fprintf(fidOrtho,'%s\n',['Monthly ' uniquePIs{i} ' ' num2str(hrs) ':' num2str(min)]);
    fprintf(fidOrtho,'%s\n',' ');
end

fprintf(fidOrtho,'%s\n',' ');
fprintf(fidOrtho,'%s\n',' ');

%Print tech time header
fprintf(fidOrtho,'%s\n','Tech time in hours');
fprintf(fidOrtho,'%s\t','Date');
fprintf(fidOrtho,'%s\t','PI');
% fprintf(fidOrtho,'%s\t','Wash U Employee Status');
fprintf(fidOrtho,'%s\t','Technician');
fprintf(fidOrtho,'%s\n','Time');

[a b] = size(userRaw);
clear userMonth;
c=0;
for i = 2:a
    if datenum(userRaw{i,1}) >= firstDate && datenum(userRaw{i,1}) < lastDate
        c=c+1;
        userMonth{c,1} = userRaw{i,1};
        userMonth{c,2} = userRaw{i,2};
        userMonth{c,3} = userRaw{i,3};
        userMonth{c,4} = userRaw{i,4};
        userMonth{c,5} = userRaw{i,5};
        userMonth{c,6} = userRaw{i,6};
        userMonth{c,7} = userRaw{i,7};
        userMonth{c,8} = userRaw{i,8};
        userMonth{c,9} = userRaw{i,9};
        if ischar(userRaw{i,8}) == 0
            if mod(userRaw{i,8},floor(userRaw{i,8})) ~= 0 && max(isnan(userRaw{i,8})) == 0
                hrs = floor(userRaw{i,8});
                min = round(mod(userRaw{i,8},1) * 60);
                if length(num2str(min)) == 2
                    userMonth{c,8} = [num2str(hrs) ':' num2str(min)];
                else
                    userMonth{c,8} = [num2str(hrs) ':0' num2str(min)];
                end
            else
                userMonth{c,8} = num2str(userRaw{i,8});
            end
        end
    end
end

fid = fopen(foiIlabs,'w');
[a b] = size(userMonth);
for i = 1:a
    if i == 1
        line = ['service_id,' 'note,' 'service_quantity,' 'purchased_on,' 'service_request_id,' 'owner_email,' 'pi_email'];
        fprintf(fid,'%s\n',line);
    end
    hrs = floor(userMonth{i,8});
    min = round(mod(userMonth{i,8},1) * 60);
   
    if strcmp(userMonth{i,7},'DEXA') == 1
        line = [...
            '306873,',...
            num2str(userMonth{i,5}), ',',...
            num2str(0.5), ',', ...
            userMonth{i,1}, ',', ...
            'n/a,',...
            userMonth{i,2}, ',', ...
            userMonth{i,3}
            ];

        fprintf(fid,'%s\n',line);
        i=i+1;

    elseif strcmp(userMonth{i,7},'X-Ray') == 1
        line = [...
                '306872,',...
                num2str(userMonth{i,5}), ',',...
                num2str(1), ',', ...
                userMonth{i,1}, ',', ...
                'n/a,',...
                userMonth{i,2}, ',', ...
                userMonth{i,3}
                ];

        fprintf(fid,'%s\n',line);

    end
end

[a b] = size(matchedInMonth);
for i = 1:a
    line = [...
            '306871,',...
            [num2str(matchedInMonth{i,4}) ':' num2str(matchedInMonth{i,5})], ',',...
            num2str(matchedInMonth{i,7}), ',', ...
            matchedInMonth{i,1}, ',', ...
            'n/a,',...
            matchedInMonth{i,6}, ',', ...
            matchedInMonth{i,2}
            ];

    fprintf(fid,'%s\n',line);
end

fclose(fid);


% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% %%Now do Micro
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% % function completeBillingCompilerMicro()
keep month firstDate endMonth lastDate
clc; 

qcNumber = 2211;

% microCTTimeCollector();%uctTime = 

%identify file of interest
foiOrtho = [['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Billing Records\'] month '2018_Micro_Billing Information.tab'];
foiIlabs = [['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Billing Records\'] month '2018_Micro_Billing Information iLabs.csv'];

%load logs
userLogPath = ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\User Logs\UserLogIlabs.txt'];
microCompleteLogPath = ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\MicroCT Machine Logs\LogFromRSQHeader*.txt'];

excel = actxserver('Excel.Application');
set(excel,'Visible',0);
workbook = excel.Workbooks;
invoke(workbook,'Open',userLogPath);
excel.ActiveWorkbook.SaveAs([pwd '\userLog.xlsx'],51);
invoke(excel, 'Quit');
delete(excel);
[userNum,userTxt,userRaw] = xlsread([pwd '\userLog.xlsx']);
sysLine = (['del ' [pwd '\userLog.xlsx']]);
system(sysLine);

ct=0;
for i = 2:length(userRaw)
    userRawDate = datenum(userRaw{i,1});
%     if userRawDate >= datenum('Jul-08-2016') && userRawDate < lastDate
        ct=ct+1;
        newUserRaw{ct,1} = userRaw{i,1};
        newUserRaw{ct,2} = userRaw{i,2};
        newUserRaw{ct,3} = userRaw{i,3};
        newUserRaw{ct,4} = userRaw{i,4};
        newUserRaw{ct,5} = userRaw{i,5};
        newUserRaw{ct,6} = userRaw{i,6};
        newUserRaw{ct,7} = userRaw{i,7};
        newUserRaw{ct,8} = userRaw{i,8};
        newUserRaw{ct,9} = userRaw{i,9};
%     end
end
userRaw = newUserRaw;



microFile = dir(microCompleteLogPath);
for i = 1:length(microFile)
    dates(i) = datenum(microFile(i).date);
end
[aa bb] = sort(dates,'descend');
microFile = microFile(bb);
microFile = microFile(1);
microRaw = importdata(['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\MicroCT Machine Logs\' microFile.name]);
[a b] = size(microRaw);
for i = 1:a
    for j = 1:b
        tmp{i,j} = microRaw(i,j);
    end
end
clear microRaw
microRaw = tmp;
clear tmp;

%pull pertinent info

[a b] = size(userRaw);
[d e] = size(microRaw);
for i = 2:a
    samples{i-1} = num2str(userRaw{i,7});
end
uniqueSamples = unique(samples);

%make a few corrections to userRaw
for i = 2:a
        if strcmpi(userRaw{i,4},'yes') == 1
            userRaw{i,4} = 'y';
        end
        if strcmpi(userRaw{i,4},'no') == 1
            userRaw{i,4} = 'n';
        end
        if strcmpi(userRaw{i,6},'yes') == 1
            userRaw{i,5} = 'y';
        end
        if strcmpi(userRaw{i,6},'no') == 1
            userRaw{i,5} = 'n';
        end
        if strcmpi(userRaw{i,8},'yes') == 1
            userRaw{i,8} = 'y';
        end
        if strcmpi(userRaw{i,8},'no') == 1
            userRaw{i,8} = 'n';
        end
        if strcmpi(userRaw{i,9},'yes') == 1
            userRaw{i,9} = 'y';
        end
        if strcmpi(userRaw{i,9},'no') == 1
            userRaw{i,9} = 'n';
        end
        userRaw{i,1} = datestr(datenum(userRaw{i,1}),1);
end

for i = 1:a
    userRaw{i,2} = upper(userRaw{i,2});
    userRaw{i,3} = upper(userRaw{i,3});
end

userRaw = userRawCorrection(userRaw);

c=0;
for i = 1:length(microRaw)
    if microRaw{i,1} >= firstDate && microRaw{i,1} < lastDate
        c=c+1;
        microMonth(c,:) = microRaw(i,:);
    end
end

samps = microMonth(:,2);
samples = unique(cell2mat(samps));
userRaw{158,5} = 501;
c=0;
for i = 1:length(samples)
    clear sampLocs userSampLocs PIs dates measurements coreMember depts users times live pilot pi s
    sampLocs = cell2mat(samps)==samples(i);
    userSampLocs = find(cell2mat(userRaw(1:end,5)) == samples(i));
    if isempty(userSampLocs)
    else
    PIs = userRaw(userSampLocs,3);
    dates = datestr(cell2mat(microMonth(sampLocs,1)));
    measurements = cell2mat(microMonth(sampLocs,3));
    coreMember = userRaw(userSampLocs,4);
    users = userRaw(userSampLocs,2);
    times = cell2mat(microMonth(sampLocs,4));
    live = userRaw(userSampLocs,6);
    
    pi = PIs{end};
    s = samples(i);
    cm = coreMember{end};
    l = live{end};
    for j = 1:length(measurements)
        c=c+1;
        matched{c,1} = dates(j,:);
        matched{c,2} = pi;
        matched{c,3} = cm;
        matched{c,4} = s;
        matched{c,5} = measurements(j);
        theUsers = unique(users);
        matched{c,6} = theUsers{end};
        matched{c,7} = times(j);
        matched{c,8} = l;
        dateNums(c) = datenum(dates(j,:));
    end
    end
end

[~,Inds] = sort(dateNums);
matched = matched(Inds,:);

matchedClean = matched;


%Pull only measurements for month of interest
count=0;
for i = 1:c
    if datenum(matchedClean{i,1}) >= firstDate && datenum(matchedClean{i,1}) < lastDate
        count=count+1;
        matchedInMonth{count,1} = matchedClean{i,1};
        matchedInMonth{count,2} = matchedClean{i,2};
        matchedInMonth{count,3} = matchedClean{i,3};
        matchedInMonth{count,4} = matchedClean{i,4};
        matchedInMonth{count,5} = matchedClean{i,5};
        matchedInMonth{count,6} = matchedClean{i,6};
        matchedInMonth{count,7} = matchedClean{i,7};
        matchedInMonth{count,8} = matchedClean{i,8};
        if matchedInMonth{count,4} == qcNumber
            matchedInMonth{count,7} = 0;
        end
            
    end
end

%arrange matchedInMonth by measurement number
measurementNumbers = cell2mat(matchedInMonth(:,5));
[B I] = sort(measurementNumbers);
matchedInMonth = matchedInMonth(I,:);

clear dates;
% [a b] = size(matchedInMonth);
% for i = 1:a
    dates = matchedInMonth(:,1);
    PIs = matchedInMonth(:,2);
% end
uniqueDates = unique(dates);
uniquePIs = unique(PIs);

[a b] = size(matchedInMonth);
dayTotal=zeros(length(uniquePIs),length(uniqueDates));
c=0;
dayUser = cell(length(uniquePIs),length(uniqueDates));
for i = 1:length(uniqueDates)
    for j = 1:length(uniquePIs)
        for k = 1:a
            if strcmp(matchedInMonth{k,1},uniqueDates{i}) == 1 && strcmp(matchedInMonth{k,2},uniquePIs{j}) == 1 
                dayTotal(j,i) = dayTotal(j,i) + matchedInMonth{k,7}; %end up with PI x day array of totals
%                 if isempty(dayUser{j,i})
%                     dayUser{j,i} = matchedInMonth{k,7};
%                 elseif ~isempty(dayUser{j,i}) && strcmpi(dayUser{j,i},matchedInMonth{k,7}) ~= 1
                    dayUser{j,i} = matchedInMonth{k,6};
%                 end
            end                
        end
    end
end

monthTotal = sum(sum(dayTotal));

 
sysLine = [['move /y "J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Billing Records\*Micro_Billing Information.tab'] '" "' ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\Micro billing log backups\"']];
system(sysLine);

fidBMD = fopen(foiOrtho, 'w');

%print headers
fprintf(fidBMD,'%s\n',['MicroCT billing report for ' month ' 2018']);
fprintf(fidBMD,'%s\n','Reference list of all scan information');
fprintf(fidBMD,'%s\t','Date');
fprintf(fidBMD,'%s\t','PI');
% fprintf(fidBMD,'%s\t','Wash U Employee Status');
fprintf(fidBMD,'%s\t','Member');
% fprintf(fidBMD,'%s\t','Dept');
fprintf(fidBMD,'%s\t','Sample');
fprintf(fidBMD,'%s\t','Measurement');
fprintf(fidBMD,'%s\t','User');
fprintf(fidBMD,'%s\t','Scan Time (minutes)');
fprintf(fidBMD,'%s\t','Scan Time (hr:mm)');
% fprintf(fidBMD,'%s\t','Live');
fprintf(fidBMD,'\n');

%print data on per-scan basis
[a b] = size(matchedInMonth);
for i = 1:a
    fprintf(fidBMD,'%s\t',matchedInMonth{i,1});
    fprintf(fidBMD,'%s\t',matchedInMonth{i,2});
    fprintf(fidBMD,'%s\t',matchedInMonth{i,3});
    fprintf(fidBMD,'%s\t',matchedInMonth{i,4});
    fprintf(fidBMD,'%s\t',num2str(matchedInMonth{i,5}));
    fprintf(fidBMD,'%s\t',num2str(matchedInMonth{i,6}));
    fprintf(fidBMD,'%s\t',[matchedInMonth{i,7}]);
    if matchedInMonth{i,7} == 0
        if matchedInMonth{i,4} == qcNumber
            fprintf(fidBMD,'%s\t',num2str(matchedInMonth{i,7}));
        end
    end
    if matchedInMonth{i,7} ~= 0
        if matchedInMonth{i,4} == qcNumber
            hrs = 0;
            min = 0;
        else
            hrs = floor(matchedInMonth{i,7});
            min = round(mod(matchedInMonth{i,7},1) * 60);
        end
        fprintf(fidBMD,'%s\t',num2str(matchedInMonth{i,7}));
        fprintf(fidBMD,'%s\t',[num2str(hrs) ':' num2str(min)]);
    end
%     fprintf(fidBMD,'%s\n',matchedInMonth{i,9});
    fprintf(fidBMD,'%s\n',' ');
end
fprintf(fidBMD,'%s\n',' ');
fprintf(fidBMD,'%s\n',' ');
fprintf(fidBMD,'%s\n','Total Period Usage');
fprintf(fidBMD,'%s\n',num2str(monthTotal));

fprintf(fidBMD,'%s\n',' ');
fprintf(fidBMD,'%s\n',' ');

% %Print daily totals per PI
% [a b] = size(dayTotal);
% fprintf(fidBMD,'%s\n','Daily totals by PI in hours:min');
% for i = 1:a
%     fprintf(fidBMD,'%s\t',uniquePIs{i});
%     for j = 1:b
%         if dayTotal(i,j) ~= 0
%             fprintf(fidBMD,'%s\t',uniqueDates{j});
%         end
%     end
%     fprintf(fidBMD,'%s\n',' ');
%     fprintf(fidBMD,'%s\t',' ');
%     for j = 1:b
%         if dayTotal(i,j) ~= 0
%             hrs = floor(dayTotal(i,j));
%             min = round(mod(dayTotal(i,j),1) * 60);
%             fprintf(fidBMD,'%s\t',[num2str(hrs) ':' num2str(min)]);
%         end
%     end
%     fprintf(fidBMD,'%s\n',' ');
%     fprintf(fidBMD,'%s\t',' ');
%     for j = 1:b
%         if dayTotal(i,j) ~= 0
% %             fprintf(fidBMD,'%s\t',dayUser{i,j});
%               fprintf(fidBMD,'%s\t',[dayUser{i,j}]);
%         end
%     end
%     fprintf(fidBMD,'%s\n',' ');
%     dailyTotal = sum(dayTotal(i,:));
%     hrs = floor(dailyTotal);
%     min = round(mod(dailyTotal,1) * 60);
%     fprintf(fidBMD,'%s\n',['Monthly ' uniquePIs{i} ' ' num2str(hrs) ':' num2str(min)]);
%     fprintf(fidBMD,'%s\n',' ');
% end
% 
% fprintf(fidBMD,'%s\n',' ');
% fprintf(fidBMD,'%s\n',' ');
% 
% %Print tech time header
% fprintf(fidBMD,'%s\n','Tech time in hours');
% fprintf(fidBMD,'%s\t','Date');
% fprintf(fidBMD,'%s\t','PI');
% % fprintf(fidBMD,'%s\t','Wash U Employee Status');
% fprintf(fidBMD,'%s\t','Member');
% fprintf(fidBMD,'%s\t','Dept');
% fprintf(fidBMD,'%s\t','Technician');
% fprintf(fidBMD,'%s\n','Time');
% 
% [a b] = size(userRaw);
% clear userMonth;
% c=0;
% for i = 2:a
%     if datenum(userRaw{i,1}) >= firstDate && datenum(userRaw{i,1}) < lastDate
%         c=c+1;
%         userMonth{c,1} = userRaw{i,1};
%         userMonth{c,2} = userRaw{i,2};
%         userMonth{c,3} = userRaw{i,3};
%         userMonth{c,4} = userRaw{i,4};
%         userMonth{c,5} = userRaw{i,5};
%         userMonth{c,6} = userRaw{i,6};
%         userMonth{c,7} = userRaw{i,7};
%         userMonth{c,8} = userRaw{i,8};
%         userMonth{c,9} = userRaw{i,9};
%         if ischar(userRaw{i,11}) == 0
%             if mod(userRaw{i,11},floor(userRaw{i,11})) ~= 0 && max(isnan(userRaw{i,11})) == 0
%                 hrs = floor(userRaw{i,11});
%                 min = round(mod(userRaw{i,11},1) * 60);
%                 if length(num2str(min)) == 2
%                     userMonth{c,11} = [num2str(hrs) ':' num2str(min)];
%                 else
%                     userMonth{c,11} = [num2str(hrs) ':0' num2str(min)];
%                 end
%             else
%                 userMonth{c,11} = num2str(userRaw{i,11});
%             end
%         end
%         
%         userMonth{c,12} = userRaw{i,12};
%         userMonth{c,13} = userRaw{i,13};
%     end
% end
% 
% 
% fprintf(fidBMD,'%s\n',' ');
% fprintf(fidBMD,'%s\n',' ');
% 
% %Print "other" header
% fprintf(fidBMD,'%s\n','Other purchases/items');
% fprintf(fidBMD,'%s\t','Date');
% fprintf(fidBMD,'%s\t','PI');
% % fprintf(fidBMD,'%s\t','Wash U Employee Status');
% fprintf(fidBMD,'%s\t','Member');
% fprintf(fidBMD,'%s\t','Dept');
% fprintf(fidBMD,'%s\t','Task/Item');
% fprintf(fidBMD,'%s\n','Quantity');

fclose('all');

%write out scans to iLabs format
%service_id note service_quantity purchased_on service_request_id
%ownder_email pi_email


fid = fopen(foiIlabs,'w');
[a b] = size(matchedInMonth);

 

for i = 1:a
    if i == 1
        line = ['service_id,' 'note,' 'service_quantity,' 'purchased_on,' 'service_request_id,' 'owner_email,' 'pi_email'];
        fprintf(fid,'%s\n',line);
    end
    hrs = floor(matchedInMonth{i,7});
    min = round(mod(matchedInMonth{i,7},1) * 60);
    endTime = addtodate(datenum(matchedInMonth{i,1}(end-7:end)),hrs,'hour');
    endTime = addtodate(endTime,min,'minute');
    
    if matchedInMonth{i,4} == qcNumber
        yesno = 'yes';
    elseif matchedInMonth{i,4} ~= qcNumber
        yesno = 'no';
    end
    
%     line = [...
%         matchedInMonth{i,7}, ', ' ,...
%         matchedInMonth{i,2}, ', ' ,...
%         'Scanco uCT 40', ', ' ,... 
%         num2str(matchedInMonth{i,5}), ', ' ,...
%         matchedInMonth{i,1}, ', ' ,...
%         datestr(endTime,'HH:MM:SS'), ', ' ,...
%         yesno...
%         ];
    line = [...
        '306870,',...
        [num2str(matchedInMonth{i,4}) ':' num2str(matchedInMonth{i,5})], ',',...
        num2str(matchedInMonth{i,7}), ',', ...
        matchedInMonth{i,1}, ',', ...
        'n/a,',...
        matchedInMonth{i,6}, ',', ...
        matchedInMonth{i,2}
        ];
        
    fprintf(fid,'%s\n',line);
%     fprintf(foiIlabs,'%s,',[matchedInMonth{i,7}]);%user
%     fprintf(foiIlabs,'%s,',matchedInMonth{i,2});%PI Last Name
%     fprintf(foiIlabs,'%s,','Scanco uCT 40');%equipment name
%     fprintf(foiIlabs,'%s,',num2str(matchedInMonth{i,5}));%sample number as project
%     fprintf(foiIlabs,'%s,',matchedInMonth{i,1});%start date and time
%     hrs = floor(matchedInMonth{i,8});
%     min = round(mod(matchedInMonth{i,8},1) * 60);
%     endTime = addtodate(datenum(matchedInMonth{i,1}(end-7:end)),hrs,'hour');
%     endTime = addtodate(endTime,min,'minute');
%     fprintf(foiIlabs,'%s,',datestr(endTime,'HH:MM:SS'));
%     if strcmpi(matchedInMonth{i,10},'y') == 1 || matchedInMonth{i,5} == qcNumber
%         fprintf(fid,'%s\n','yes');
%     elseif strcmpi(matchedInMonth{i,10},'y') == 0 && matchedInMonth{i,5} ~= qcNumber
%         fprintf(fid,'%s\n','no');
%     end
end
fclose(fid);
    
function [userRaw] = userRawCorrection(userRaw)

%PIs
for i = 1:length(userRaw)
    if ~isempty(strfind(userRaw{i,3},'SCHELLER'))
        userRaw{i,3} = 'escheller@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'TANG'))
        userRaw{i,3} = 'tangs@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'O''KEEFE'))
        userRaw{i,3} = 'okeefer@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'OKEEFE'))
        userRaw{i,3} = 'okeefer@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'SILVA'))
        userRaw{i,3} = 'silvam@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'TEITELBAUM'))
        userRaw{i,3} = 'teitelbs@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'YOSEF'))
        userRaw{i,3} = 'abuamery@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'AMER'))
        userRaw{i,3} = 'abuamery@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'ABU-AMER'))
        userRaw{i,3} = 'abuamery@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'CRAFT'))
        userRaw{i,3} = 'clarissa.craft@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'PETERSON'))
        userRaw{i,3} = 'timrpeterson@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'MBALAVIELE'))
        userRaw{i,3} = 'gmbalaviele@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'SANDELL'))
        userRaw{i,3} = 'sandelll@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'LONGF'))
        userRaw{i,3} = 'flong@wustl.edu';
    end
    if strcmp(userRaw{i,3},'LONG') == 1
        userRaw{i,3} = 'flong@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'FANXIN'))
        userRaw{i,3} = 'flong@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'F. Long'))
        userRaw{i,3} = 'flong@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'STEWART'))
        userRaw{i,3} = 'sheila.stewart@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'NEPPLE'))
        userRaw{i,3} = 'nepplej@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'FACCIO'))
        userRaw{i,3} = 'faccior@wustl.edu';
    end 
    if ~isempty(strfind(userRaw{i,3},'ORNITZ'))
        userRaw{i,3} = 'dornitz@wustl.edu';
    end 
    if ~isempty(strfind(userRaw{i,3},'NOVACK'))
        userRaw{i,3} = 'novack@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'MECHAM'))
        userRaw{i,3} = 'bmecham@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'SHEN'))
        userRaw{i,3} = 'hshen22@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'MCALINDEN'))
        userRaw{i,3} = 'mcalindena@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'CIVITELLI'))
        userRaw{i,3} = 'civitellir@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'WEILBAECHER'))
        userRaw{i,3} = 'kweilbae@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'ARBEIT'))
        userRaw{i,3} = 'arbeitj@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'REMEDI'))
        userRaw{i,3} = 'mremedi@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'HARRIS'))
        userRaw{i,3} = 'harrisc@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'NEPPLE'))
        userRaw{i,3} = 'nepplej@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'nepple'))
        userRaw{i,3} = 'nepplej@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,3},'debabrata'))
        userRaw{i,3} = 'patrad@wudosis.wustl.edu';
    end
    
    
    
    
end

%users
for i = 1:length(userRaw)
    userRaw{i,2} = upper(userRaw{i,2});
    if ~isempty(strfind(userRaw{i,2},'BUETTMAN'))
        userRaw{i,2} = 'buettmannev@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'COATES'))
        userRaw{i,2} = 'coatesb@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'ZOU'))
        userRaw{i,2} = 'weizou@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'KANEKO'))
        userRaw{i,2} = 'keikokaneko@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'KEIKO'))
        userRaw{i,2} = 'keikokaneko@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'PEI'))
        userRaw{i,2} = 'phu22@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'HUP@'))
        userRaw{i,2} = 'phu22@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'DEYE'))
        userRaw{i,2} = 'songda@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'BOER'))
        userRaw{i,2} = 'libo@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'MIGOTSKY'))
        userRaw{i,2} = 'n.migotsky@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'COWARDIN'))
        userRaw{i,2} = 'CCOWARDIN@path.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'KARLY.LORBEER@WUSTL.EDU'))
        userRaw{i,2} = 'okeefr@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'HONG'))
        userRaw{i,2} = 'chenho@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'MANOJ'))
        userRaw{i,2} = 'arram@wusm.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'SEUNG-YON'))
        userRaw{i,2} = 'seung-yonlee@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'MANUELA FORTUNATO'))
        userRaw{i,2} = 'manuela.fortunato@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'TAOTAO'))
        userRaw{i,2} = 'taotao.xu@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'LIPNER'))
        userRaw{i,2} = 'lipnerj@wudosis.wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'FONTANA'))
        userRaw{i,2} = 'FRANCESCA.FONTANA@WUSTL.EDU';
    end
    if ~isempty(strfind(userRaw{i,2},'YAEL'))
        userRaw{i,2} = 'yalippe@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'JAYARAM'))
        userRaw{i,2} = 'rohith.jayaram@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'GAURAV'))
        userRaw{i,2} = 'gauravswarnkar@wustl.edu';
    end
    if ~isempty(strfind(userRaw{i,2},'swarnkar'))
        userRaw{i,2} = 'gauravswarnkar@wustl.edu';
    end
    
    if ~isempty(strfind(userRaw{i,2},'XUT'))
        userRaw{i,2} = 'taotao.xu@wustl.edu';
    end
    
    if ~isempty(strfind(userRaw{i,2},'YU SHI'))
        userRaw{i,2} = 'yushi@wustl.edu';
    end
    
    
    
    
    
    
end
end