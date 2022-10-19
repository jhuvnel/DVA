function MVI_path = add_to_tmDVAfile(MVI_path)
if nargin < 1 || isempty(MVI_path)
    prompt = 'Select the MVI Study subject root folder.';
    MVI_path = uigetdir(prompt,prompt);
    if ~contains(MVI_path,'MVI')
        disp(['The selected path does not contain the text "MVI", so it may be wrong: ',MVI_path])
    end
end
short_fnames = extractfield(dir([MVI_path,filesep,'MVI*R*',filesep,'MVI*_tmDVA.xlsx']),'name');
MVI_fnames = strcat(extractfield(dir([MVI_path,filesep,'MVI*R*',filesep,'MVI*_tmDVA.xlsx']),'folder'),filesep,short_fnames);
opts = [short_fnames;{'Make New'}];
tf = 1;
resp = '';
prompt = {['Add data for subject.',newline,'Leave blank if not attempted.',newline,newline,'Subject: '];...
    'Visit# (ex.11x): ';'Date (mm/dd/yy): ';'Condition: ';'VA at 0mph: ';...
    'VA at 0.5mph: ';'VA at 1.0mph: ';'VA at 1.5mph: ';'VA at 2.0mph: ';...
    'VA at 2.5mph: ';'VA at 3.0mph: '};
def = repmat({''},length(prompt),1); %Default prompt    
empty_warnings = {'Missing subject name.','Missing visit name.','Missing date.','Missing condition name.','Missing SVA data.'};
while tf %Keep running until the user selects to cancel
    if contains(resp,short_fnames) %File already exists and will be appended to
        %Load file
        [~,~,scores] = xlsread(MVI_fnames{ind});
        %Remove erroneous lines at the end (NaN in all values)
        scores(cellfun(@isnumeric,scores(:,1)),:) = [];
        %Make tab
        old_tab = [cell2table(scores(2:end,1:4),'VariableNames',scores(1,1:4)),array2table(cell2mat(scores(2:end,5:8)),'VariableNames',scores(1,5:8))];
        %Turn date into a datetime array
        old_tab.Date = datetime(old_tab.Date);
        %Unique Entries
        present_already = unique(strcat({'Visit'},scores(2:end,2),{' '},scores(2:end,3),{' '},scores(2:end,4)),'stable');
        %Display which values already exist
        disp(['Already present in ',MVI_fnames{ind}])
        disp(strjoin(present_already,newline))
        valid_info = false; %Run until true
        def{1} = strrep(resp,'_tmDVA.xlsx','');
        while(~valid_info)
            info = inputdlg(prompt,'Input tmDVA data to be added to file',1,def);
            if isempty(info) %Return if user hit cancel here
                return;
            end
            def = info; %Set the previous response as the default in case you need to run again
            empty_cells = cellfun(@isempty,info);
            is_numeric = ~isnan(str2double(info));
            date = info{3};
            %Error handle invalid inputs
            if any(empty_cells(1:5)) 
                %Missing subject, visit, date, condition, or SVA (which everyone should do, other values can be blank)
                disp(strjoin(empty_warnings(empty_cells(1:5)),newline))
            elseif length(date)~= 8 || ~all(ismember(date([1,2,4,5,7,8]),'0123456789')) || ~all(ismember(date([3,6]),'/-'))
                disp('Date entered is not in the format mm/dd/yy. Try again.')    
            elseif any(~empty_cells(5:end)&~is_numeric(5:end))
                disp('Visual acuity values must be numeric or left blank.')
            else
                valid_info = true;
                def(5:end) = {''};
            end
        end
        sub = info{1};
        vis = strrep(strrep(info{2},' ',''),'Visit','');
        date = datetime(info{3},'InputFormat','MM/dd/yy');
        cond = strrep(info{4},' ','');
        DVA_dat = str2double(info(5:end));
        %Make the new tab
        new_tab = table(); 
        new_tab.Subject = repmat({sub},7,1);
        new_tab.Visit = repmat({vis},7,1);
        new_tab.Date = repmat(date,7,1);
        new_tab.Condition = repmat({cond},7,1);
        new_tab.Speed = [0;0.5;1.0;1.5;2.0;2.5;3.0];
        new_tab.SVA = DVA_dat(1)*ones(7,1);
        new_tab.DVA = DVA_dat;
        new_tab.('SVA-DVA') = DVA_dat(1)-DVA_dat;
        if any(contains(old_tab.Condition(old_tab.Date==date),cond)) %Already exists in the folder
            disp(['Visit ',vis,' and conditon ',cond,' already exists in the Excel file for subject ',resp,' and is being overwritten.'])
            tab = old_tab;
            tab(contains(old_tab.Condition(old_tab.Date==date),cond),:) = new_tab;
        else
            tab = [old_tab;new_tab];
        end
        tab = sortrows(tab,'Date','Ascend');
        writetable(tab,MVI_fnames{ind})
        disp(['Saved to ',MVI_fnames{ind}])
    elseif contains(resp,'Make New')
        out_path = uigetdir('Select where to save this file.','Select where to save this file.');
        if isnumeric(out_path) %Return if user hit cancel here
            return;
        end
        valid_info = false; %Run until true
        while(~valid_info)
            info = inputdlg(prompt,'Input tmDVA data to be added to file',1,def);
            if isempty(info) %Return if user hit cancel here
                return;
            end
            def = info; %Set the previous response as the default in case you need to run again
            empty_cells = cellfun(@isempty,info);
            is_numeric = ~isnan(str2double(info));
            date = info{3};
            %Error handle invalid inputs
            if any(empty_cells(1:5)) 
                %Missing subject, visit, date, condition, or SVA (which everyone should do, other values can be blank)
                disp(strjoin(empty_warnings(empty_cells(1:5)),newline))
            elseif length(date)~= 8 || ~all(ismember(date([1,2,4,5,7,8]),'0123456789')) || ~all(ismember(date([3,6]),'/-'))
                disp('Date entered is not in the format mm/dd/yy. Try again.')    
            elseif any(~empty_cells(5:end)&~is_numeric(5:end))
                disp('Visual acuity values must be numeric or left blank.')
            else
                valid_info = true;
                def(5:end) = {''};
            end
        end
        sub = info{1};
        vis = strrep(strrep(info{2},' ',''),'Visit','');
        date = datetime(info{3},'InputFormat','MM/dd/yy');
        cond = strrep(info{4},' ','');
        DVA_dat = str2double(info(5:end));
        %Make the new tab
        tab = table(); 
        tab.Subject = repmat({sub},7,1);
        tab.Visit = repmat({vis},7,1);
        tab.Date = repmat(date,7,1);
        tab.Condition = repmat({cond},7,1);
        tab.Speed = [0;0.5;1.0;1.5;2.0;2.5;3.0];
        tab.SVA = DVA_dat(1)*ones(7,1);
        tab.DVA = DVA_dat;
        tab.('SVA-DVA') = DVA_dat(1)-DVA_dat;
        out_file = [out_path,filesep,sub,'-tmDVA.xlsx'];
        writetable(tab,out_file)
        disp(['Saved to ',out_file])
    end    
    % Poll for new reponse
    [ind,tf] = listdlg('PromptString','Select a file to add to:','SelectionMode','single',...
        'ListSize',[200 200],'ListString',opts);
    if tf
        resp = opts{ind};
    end    
end
end