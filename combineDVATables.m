function [all_tmDVA,MVI_path] = combineDVATables(MVI_path)
if nargin < 1 || isempty(MVI_path)
    prompt = 'Select the MVI Study subject root folder.';
    MVI_path = uigetdir(prompt,prompt);
    if ~contains(MVI_path,'MVI')
        disp(['The selected path does not contain the text "MVI", so it may be wrong: ',MVI_path])
    end
end
MVI_fnames = strcat(extractfield(dir([MVI_path,filesep,'MVI*R*',filesep,'MVI*_tmDVA.xlsx']),'folder'),filesep,extractfield(dir([MVI_path,filesep,'MVI*R*',filesep,'MVI*_tmDVA.xlsx']),'name'));
all_tmDVA = cell(length(MVI_fnames),1);
%Assumes they all have the same score order
for i = 1:length(MVI_fnames)
    %Reads the first sheet as the right sheet, thankfully
    [~,~,scores] = xlsread(MVI_fnames{i});
    %Remove erroneous lines at the end (NaN in all values)
    scores(cellfun(@isnumeric,scores(:,1)),:) = [];
    %Make tab
    tab = [cell2table(scores(2:end,1:4),'VariableNames',scores(1,1:4)),array2table(cell2mat(scores(2:end,5:8)),'VariableNames',scores(1,5:8))];
    %Turn date into a datetime array
    tab.Date = datetime(tab.Date);
    all_tmDVA(i) = {tab};
end
all_tmDVA = vertcat(all_tmDVA{:});
save([MVI_path,filesep,'ALLMVI-tmDVA.mat'],'all_tmDVA')
writetable(all_tmDVA,[MVI_path,filesep,'ALLMVI-tmDVA.xlsx'])
end