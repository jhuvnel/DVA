%% DVAA.m
% Dynamic Visual Acuity Analyzer
%
%A list-based script that allowed for using the full functionality of the
%rest of the scripts/functions in this repository.
%MDA Menu Options
opts = {'Add Experiment to File','Make Summary File',...
    'Plot MVI Summary','Plot Candidate'};
resp1 = '';
tf1 = 1;
MVI_path = '';
while tf1
    switch resp1
        case 'Add Experiment to File'
            MVI_path = add_to_tmDVAfile(MVI_path);
        case 'Make Summary File'
            [all_tmDVA,MVI_path] = combineDVATables(MVI_path);
        case  'Plot MVI Summary'
            
        case 'Plot Candidate'

    end
    % Poll for new reponse
    [ind1,tf1] = listdlg('PromptString','Select an action:','SelectionMode','single',...
                       'ListSize',[150 200],'ListString',opts); 
    if tf1
        resp1 = opts{ind1}; 
    end
end
disp('QOLA instance ended.')