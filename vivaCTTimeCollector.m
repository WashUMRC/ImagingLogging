function [outMat] = vivaCTTimeCollector()

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%This function acts as a collector for scan time data for the main
%%function           . It executes a command file on the remote Scanco
%%server that exports RSQ header data that will be read in for time
%%reporting purposes.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%RSQHEADERTEMPLATE.COM text:
%
% $ DEFINE/user sys$output dk0:[microct.data.SAMPLE.MEASUREMENT]RSQHeader1.txt
% $ ctheader dk0:[microct.data.SAMPLE.MEASUREMENT]CNUMBER.RSQ;1
% $ EXIT
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

templatePath = [pwd '\RSQHEADERTEMPLATE.COM'];
comFilePath = [pwd '\microCTComFile.com'];
serverIP = '10.21.24.203';
remoteScratch = 'IDISK1:[MICROCT.SCRATCH]';

plinkPath = '"C:\Program Files\PuTTY\plink.exe" ';
savedSession = 'VivaCT40 ';
userName = 'microct ';
password = 'mousebone4 ';

savePath = '"J:\Silva''s Lab\P30 Core Center\SSC_Billing Information\Viva RSQ Header Text Files\"';
theDate = datestr(now);
logFileOut = ['J:\Silva''s Lab\P30 Core Center\Core B_Billing Information\VivaCT Machine Logs\LogFromRSQHeader' theDate(1:11) '.txt'];

%read template for generating RSQ header text file
template = importdata(templatePath);

%generate com file to be used to generate headers files
%to run the com file, you must first execute SET FILE/ATTRIBUTE=RFM:STM and
%type in the full path to the com file when prompted.
comFile = fopen(comFilePath,'wt');
fprintf(comFile,'%s','$ SET FILE/ATTRIBUTEs=RFM=STM');
fprintf(comFile,'%s\n','');
fprintf(comFile,'%s\n','$! Original code created by Dan Leib');
f = ftp(serverIP,'microct','mousebone4');
ascii(f);
cd(f,'dk0');
cd(f,'data');
dirs = dir(f);%samples
for i = 1:length(dirs)
    if dirs(i).isdir == 1%check if is a directory
        cd(f,dirs(i).name(1:(end-2)));
        dirs2 = dir(f);%measurements
        for j = 1:length(dirs2)
            if dirs2(j).isdir == 1
                cd(f,dirs2(j).name(1:end-2));
                rsqs = dir(f,'*.rsq*');
                for k = 1:length(rsqs)
                    line1 = strrep(template{1},'SAMPLE',dirs(i).name(1:end-6));
                    line1 = strrep(line1,'MEASUREMENT',dirs2(j).name(1:end-6));
                    line2 = strrep(template{2},'SAMPLE',dirs(i).name(1:end-6));
                    line2 = strrep(line2,'MEASUREMENT',dirs2(j).name(1:end-6));
                    line2 = strrep(line2,'CNUMBER',rsqs(k).name(1:end-6));
                    fprintf(comFile,'%s\n',line1);
                    fprintf(comFile,'%s\n',line2);
                end
                cd(f,'..');%back out of measurement
            end
        end
        cd(f,'..');%back out of sample
    end
end
fprintf(comFile,'%s\n','$ logoff');
fclose(comFile);
cd(f,remoteScratch);
mput(f,comFilePath);

sysLine = [plinkPath savedSession '-l ' userName '-pw ' password '@' remoteScratch 'microctcomfile.com'];
system(sysLine);

    
%get all the RSQ header files
cd(f,'dk0:[microct.data]');
for i = 1:length(dirs)
    if dirs(i).isdir == 1%check if is a directory
        cd(f,dirs(i).name(1:(end-2)));
        dirs2 = dir(f);%measurements
        for j = 1:length(dirs2)
            if dirs2(j).isdir == 1
                cd(f,dirs2(j).name(1:end-2));
                rsqHeaders = dir(f,'RSQHEADER*.TXT*');
                if ~isempty(rsqHeaders)
                    num = length(rsqHeaders);
                    sysLine = ['md "' savePath(2:end-1) dirs(i).name(1:end-6) '\' dirs2(j).name(1:end-6) '"' ];
                    system(sysLine);
                    mget(f,rsqHeaders(num).name,[savePath(2:end-1) dirs(i).name(1:end-6) '\' dirs2(j).name(1:end-6)]);
                    delete(f,rsqHeaders(num).name);
                end
                 cd(f,'..');%back out of measurement
%                  f
            end
        end
        cd(f,'..');%back out of sample
    end
end
                    
%Calculate scan times based on RSQ header information
c=0;
dirs = dir(savePath(2:end-1));
for i = 3:length(dirs)
    if dirs(i).isdir == 1
        dirs2 = dir([savePath(2:end-1) dirs(i).name]);
        for j = 3:length(dirs2)
            if dirs2(j).isdir == 1
                headerFiles = dir([savePath(2:end-1) dirs(i).name '\' dirs2(j).name '\*.txt*']);
                for k = length(headerFiles)
                    fid = fopen([savePath(2:end-1) dirs(i).name '\' dirs2(j).name '\' headerFiles(k).name]);
                    ct=0;
                    while ~feof(fid)
                    	ct = ct+1;
                        line{ct} = fgets(fid);
                    end
                    line = line';
                    dvLine = strfind(line,'# Detectors V :');
                    dBinVLine = strfind(line,'Detector Binning V');
                    intTimeLine = strfind(line,'Integration-Time');
                    dimZLine = strfind(line,'Dim Z');
                    projStackLine = strfind(line,'No frames/stack');
                    for l = 1:ct
                        if ~isempty(dvLine{l})
                            dV = str2num(line{l}(25:29));
                        end
                        if ~isempty(dBinVLine{l})
                            dbinV = str2num(line{l}(23:30));
                        end
                        if ~isempty(dimZLine{l})
                            dimz = str2num(line{l}(22:29));
                        end
                        if ~isempty(projStackLine{l})
                             numProjectionsstack = str2num(line{l}(22:29));
                        end
                        if ~isempty(intTimeLine{l})
                             integrationtime = str2num(line{l}(22:29)) / 1000;%in ms
                        end
                    end
                    date = line{11}(22:end);
                    readouttime = 110;%317; %constant for microct   
                    numStacks = ceil(dimz / (dV / dbinV));
                    time = numStacks * numProjectionsstack * (integrationtime + readouttime);
                    time = time / 1000';
                    time = time / 60;
                    time = time / 60; %ends up in hours
                    
                    %put these things in an output matrix
                    c=c+1;
                    outMat(c,1) = datenum(date);
                    outMat(c,3) = str2num(dirs2(j).name);
                    outMat(c,4) = time;
                    outMat(c,2) = str2num(dirs(i).name);
                    fclose(fid);
                end
            end
        end
    end
end


fileOut = logFileOut;
fid = fopen(fileOut,'w');
[a b] = size(outMat);
for i = 1:a
    for j = 1:b
        if j ~= b
            fprintf(fid,'%s\t',num2str(outMat(i,j)));
        else
            fprintf(fid,'%s\n',num2str(outMat(i,j)));
        end
    end
end
fclose(fid);

% exit;
