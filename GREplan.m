nList = 31;
ended = 0;
iday = 1;
iList = 1;
Plan = cell(floor(nList/2)+1+15,1);

while ~ended
    Plan{iday} = [Plan{iday} 'New list:' num2str(iList) '   '];
    Plan{iday+2} = [Plan{iday+2} 'Review List:' num2str(iList) '   '];
    Plan{iday+4} = [Plan{iday+4} 'Review List:' num2str(iList) '   '];
    Plan{iday+7} = [Plan{iday+7} 'Review List:' num2str(iList) '   '];
    Plan{iday+15} = [Plan{iday+15} 'Review List:' num2str(iList) '   '];
    iList = iList +1;
    iday = floor(iList/2)+1;
    if iList > nList
       ended = 1;
    end
end

xlswrite('GRE.xls',Plan)
