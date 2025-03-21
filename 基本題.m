%% ========== 基本初始化 ========== 
clc; clear; close all;

% 設定檔案路徑 (依實際需要修改)
filePath = 'C:\Users\user\Desktop\113-2\TSMC WS\2025workshop_若選擇Part A_Scheduling題目需要使用.xlsx';

% 讀取資料 (共 40 筆 WIP，故讀取範圍 2:41 共 40 rows)
ini_cart_lo = readmatrix(filePath,'Sheet','CART','Range','B2:B21');           % 推車初始位置 (20 筆)
remaining_Q = readmatrix(filePath,'Sheet','WIP','Range','B2:B41');           % WIP 剩餘 Q-Time (40 筆)
wip_F       = readmatrix(filePath,'Sheet','WIP','Range','C2:C41');           % WIP 取貨位置 (40 筆)
wip_T       = readmatrix(filePath,'Sheet','WIP','Range','D2:D41');           % WIP 送達位置 (40 筆)
dis_matrix  = readmatrix(filePath,'Sheet','Distance_Matrix','Range','B2:AY51'); % 距離矩陣
wip_ids     = readcell(filePath, 'Sheet', 'WIP', 'Range', 'A2:A41');         % WIP ID (40 筆)

num_cart = 20;
num_wip  = 40;  % WIP 總筆數

%% ========== 建立所有可能路由 ========== 
% 結構包含：wips, strategy, arriveTimes, routeDist, feasibleCart
route_list = struct('wips',{}, 'strategy',{}, 'arriveTimes',{}, 'routeDist',[], 'feasibleCart',[]);

r_idx = 0;
% --- 雙 WIP 路由 (PD-PD 與 PP-DD) --- 
for w1 = 1:(num_wip-1)
    for w2 = (w1+1):num_wip
        % PD-PD 路由
        r_idx = r_idx + 1;
        route_list(r_idx).wips         = [w1, w2];
        route_list(r_idx).strategy     = 'PD-PD';
        route_list(r_idx).arriveTimes  = cell(num_cart,1);
        route_list(r_idx).routeDist    = zeros(num_cart,1);
        route_list(r_idx).feasibleCart = false(num_cart,1);
        
        % PP-DD 路由
        r_idx = r_idx + 1;
        route_list(r_idx).wips         = [w1, w2];
        route_list(r_idx).strategy     = 'PP-DD';
        route_list(r_idx).arriveTimes  = cell(num_cart,1);
        route_list(r_idx).routeDist    = zeros(num_cart,1);
        route_list(r_idx).feasibleCart = false(num_cart,1);
    end
end
num_routes = length(route_list);

%% ========== 計算每條路由對每台車是否可行及其距離 ========== 
for r = 1:num_routes
    wipsR = route_list(r).wips;
    strat = route_list(r).strategy;
    for c = 1:num_cart
        initPos = ini_cart_lo(c);
        % 呼叫輔助函式計算到達時間、總距離與可行性
        [arrTimes, totDist, isFeasible] = computeRouteMetrics(wipsR, strat, initPos, wip_F, wip_T, dis_matrix, remaining_Q);
        route_list(r).arriveTimes{c}  = arrTimes;
        route_list(r).routeDist(c)    = totDist;
        route_list(r).feasibleCart(c) = isFeasible;
    end
end

%% ========== (3) 第一階段：最小化報廢數 (delta) ========== 
prob1 = optimproblem;

% 決策變數
y     = optimvar('y', [num_cart, num_routes], 'Type','integer','LowerBound',0,'UpperBound',1);
delta = optimvar('delta', [num_wip,1],        'Type','integer','LowerBound',0,'UpperBound',1);

% 目標：最小化報廢數
prob1.Objective = sum(delta);

% (A) 每台車最多選 1 條路由
conCar = optimconstr(num_cart,1);  
for c = 1:num_cart
    conCar(c) = sum(y(c,:)) <= 1;  
end
prob1.Constraints.conCar = conCar; 

% (B) 不可行路由：y(c,r)=0
% 先計算不可行的 (c,r) 數量
num_infeasible = sum(~[route_list.feasibleCart], 'all');  
conFeasible    = optimconstr(num_infeasible,1);

idx = 1;
for r = 1:num_routes
    for c = 1:num_cart
        if ~route_list(r).feasibleCart(c)
            conFeasible(idx) = (y(c,r) == 0);
            idx = idx + 1;
        end
    end
end
if num_infeasible > 0
    prob1.Constraints.conFeasible = conFeasible;
end

% (C) 每個 WIP 只能被 1 條路由服務，否則報廢 => sum(...) + delta(w) = 1
for w = 1:num_wip
    relevantExpr = 0;
    found = false;
    for r = 1:num_routes
        if ismember(w, route_list(r).wips)
            for c = 1:num_cart
                if route_list(r).feasibleCart(c)
                    relevantExpr = relevantExpr + y(c,r);
                    found = true;
                end
            end
        end
    end
    if found
        prob1.Constraints.(['conWIP_' num2str(w)]) = (relevantExpr + delta(w) == 1);
    else
        % 如果沒有任何車可行，該 WIP 必須報廢
        prob1.Constraints.(['conWIP_' num2str(w)]) = (delta(w) == 1);
    end
end

opts  = optimoptions('intlinprog','Display','final','Heuristics','advanced');
sol1  = solve(prob1, 'Options', opts);
minScrap = sum(sol1.delta);
fprintf('第一階段: 最小報廢數 = %d\n', minScrap);

% --- Proceed with further steps for minimizing total distance ---
% The rest of the code follows as per the problem structure


%% ========== (4) 第二階段：在固定報廢數下最小化總距離 ==========
prob2 = optimproblem;
y2    = optimvar('y2', [num_cart, num_routes], 'Type','integer','LowerBound',0,'UpperBound',1);
delta2= optimvar('delta2', [num_wip,1],        'Type','integer','LowerBound',0,'UpperBound',1);

% 目標：最小化車隊總距離
totalDistExpr = 0;
for c = 1:num_cart
    for r = 1:num_routes
        totalDistExpr = totalDistExpr + route_list(r).routeDist(c)*y2(c,r);
    end
end
prob2.Objective = totalDistExpr;

% (A) 固定報廢數
prob2.Constraints.fixScrap = sum(delta2) == minScrap;

% (B) 每台車最多選1條路由
conCar2 = optimconstr(num_cart,1);  
for c = 1:num_cart
    conCar2(c) = sum(y2(c,:)) <= 1;  
end
prob2.Constraints.conCar2 = conCar2; 

% (C) 不可行路由 => y2(c,r)=0
num_infeasible2 = sum(~[route_list.feasibleCart], 'all');
conFeasible2    = optimconstr(num_infeasible2,1);

idx2 = 1;
for r = 1:num_routes
    for c = 1:num_cart
        if ~route_list(r).feasibleCart(c)
            conFeasible2(idx2) = (y2(c,r)==0);
            idx2 = idx2 + 1;
        end
    end
end
if num_infeasible2 > 0
    prob2.Constraints.conFeasible2 = conFeasible2;
end

% (D) WIP 指派或報廢 => sum(...) + delta2(w) = 1
for w = 1:num_wip
    relevantExpr2 = 0;
    found = false;
    for r = 1:num_routes
        if ismember(w, route_list(r).wips)
            for c = 1:num_cart
                if route_list(r).feasibleCart(c)
                    relevantExpr2 = relevantExpr2 + y2(c,r);
                    found = true;
                end
            end
        end
    end
    if found
        prob2.Constraints.(['conWIP2_' num2str(w)]) = (relevantExpr2 + delta2(w) == 1);
    else
        prob2.Constraints.(['conWIP2_' num2str(w)]) = (delta2(w) == 1);
    end
end

sol2 = solve(prob2, 'Options', opts);
finalScrap = sum(sol2.delta2);
finalDist  = 0;
for c = 1:num_cart
    for r = 1:num_routes
        finalDist = finalDist + route_list(r).routeDist(c)*sol2.y2(c,r);
    end
end
fprintf('第二階段: 報廢數固定 = %d, 總距離 = %.2f\n', finalScrap, finalDist);

%% ========== (5) 產生排程明細 ==========
% 輸出格式：CART_ID、ORDER、WIP_ID、ACTION、COMPLETE_TIME
scheduleData = {};  
rowCount     = 0;       

for c = 1:num_cart
    chosenR = find(sol2.y2(c,:) > 0.5);
    if isempty(chosenR), continue; end
    for rr = chosenR
        wipsUsed  = route_list(rr).wips;
        stratUsed = route_list(rr).strategy;
        initPos   = ini_cart_lo(c);
        
        if strcmp(stratUsed, 'SINGLE')
            % 單一 WIP 路由
            w = wipsUsed(1);
            step1 = dis_matrix(initPos, wip_F(w));
            step2 = dis_matrix(wip_F(w), wip_T(w));
            completeTime1 = step1;              % PICKUP 完成時間
            completeTime2 = step1 + step2;      % DELIVERY 完成時間
            
            rowCount = rowCount + 1;
            scheduleData{rowCount,1} = c;       % CART_ID
            scheduleData{rowCount,2} = 1;       % ORDER
            scheduleData{rowCount,3} = wip_ids{w};
            scheduleData{rowCount,4} = 'PICKUP';
            scheduleData{rowCount,5} = completeTime1;
            
            rowCount = rowCount + 1;
            scheduleData{rowCount,1} = c;
            scheduleData{rowCount,2} = 2;
            scheduleData{rowCount,3} = wip_ids{w};
            scheduleData{rowCount,4} = 'DELIVERY';
            scheduleData{rowCount,5} = completeTime2;
            
        else
            % 雙 WIP 路由
            if strcmp(stratUsed, 'PD-PD')
                w1 = wipsUsed(1); 
                w2 = wipsUsed(2);
                step1 = dis_matrix(initPos, wip_F(w1));
                step2 = dis_matrix(wip_F(w1), wip_T(w1));
                completeTime1 = step1;
                completeTime2 = step1 + step2;
                
                step3 = dis_matrix(wip_T(w1), wip_F(w2));
                step4 = dis_matrix(wip_F(w2), wip_T(w2));
                completeTime3 = completeTime2 + step3;
                completeTime4 = completeTime3 + step4;
                
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 1;
                scheduleData{rowCount,3} = wip_ids{w1};
                scheduleData{rowCount,4} = 'PICKUP';
                scheduleData{rowCount,5} = completeTime1;
                
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 2;
                scheduleData{rowCount,3} = wip_ids{w1};
                scheduleData{rowCount,4} = 'DELIVERY';
                scheduleData{rowCount,5} = completeTime2;
                
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 3;
                scheduleData{rowCount,3} = wip_ids{w2};
                scheduleData{rowCount,4} = 'PICKUP';
                scheduleData{rowCount,5} = completeTime3;
                
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 4;
                scheduleData{rowCount,3} = wip_ids{w2};
                scheduleData{rowCount,4} = 'DELIVERY';
                scheduleData{rowCount,5} = completeTime4;
                
            elseif strcmp(stratUsed, 'PP-DD')
                w1 = wipsUsed(1); 
                w2 = wipsUsed(2);
                step1 = dis_matrix(initPos, wip_F(w1));
                step2 = dis_matrix(wip_F(w1), wip_F(w2));
                step3 = dis_matrix(wip_F(w2), wip_T(w1));
                completeTime1 = step1 + step2 + step3;
                
                step4 = dis_matrix(wip_T(w1), wip_T(w2));
                completeTime2 = completeTime1 + step4;
                
                % PICKUP w1
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 1;
                scheduleData{rowCount,3} = wip_ids{w1};
                scheduleData{rowCount,4} = 'PICKUP';
                scheduleData{rowCount,5} = dis_matrix(initPos, wip_F(w1));
                
                % PICKUP w2
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 2;
                scheduleData{rowCount,3} = wip_ids{w2};
                scheduleData{rowCount,4} = 'PICKUP';
                scheduleData{rowCount,5} = ...
                    dis_matrix(initPos, wip_F(w1)) + dis_matrix(wip_F(w1), wip_F(w2));
                
                % DELIVERY w1
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 3;
                scheduleData{rowCount,3} = wip_ids{w1};
                scheduleData{rowCount,4} = 'DELIVERY';
                scheduleData{rowCount,5} = completeTime1;
                
                % DELIVERY w2
                rowCount = rowCount + 1;
                scheduleData{rowCount,1} = c;
                scheduleData{rowCount,2} = 4;
                scheduleData{rowCount,3} = wip_ids{w2};
                scheduleData{rowCount,4} = 'DELIVERY';
                scheduleData{rowCount,5} = completeTime2;
            end
        end
    end
end

% 轉為 table 格式，並於命令視窗顯示
scheduleDataTable = cell2table(scheduleData, ...
    'VariableNames',{'CART_ID','ORDER','WIP_ID','ACTION','COMPLETE_TIME'});
disp(scheduleDataTable);

%% ========== 輔助函式：計算路由到達時間、總距離與可行性 ==========
function [arrTimes, totDist, isFeasible] = computeRouteMetrics(wips, strategy, initPos, wip_F, wip_T, dis_matrix, remaining_Q)
    if numel(wips) == 1
        % 單一 WIP 路由
        w = wips;
        step1    = dis_matrix(initPos, wip_F(w));
        step2    = dis_matrix(wip_F(w), wip_T(w));
        arrTimes = step1 + step2;
        totDist  = arrTimes;
        isFeasible = (arrTimes <= remaining_Q(w));
        
    else
        % 雙 WIP 路由
        w1 = wips(1); 
        w2 = wips(2);
        
        switch strategy
            case 'PD-PD'
                step1 = dis_matrix(initPos,  wip_F(w1));
                step2 = dis_matrix(wip_F(w1), wip_T(w1));
                arrW1 = step1 + step2;
                
                step3 = dis_matrix(wip_T(w1), wip_F(w2));
                step4 = dis_matrix(wip_F(w2), wip_T(w2));
                arrW2 = arrW1 + step3 + step4;
                
                totDist   = step1 + step2 + step3 + step4;
                arrTimes  = [arrW1, arrW2];
                isFeasible = (arrW1 <= remaining_Q(w1)) && (arrW2 <= remaining_Q(w2));
                
            case 'PP-DD'
                step1 = dis_matrix(initPos,   wip_F(w1));
                step2 = dis_matrix(wip_F(w1), wip_F(w2));
                step3 = dis_matrix(wip_F(w2), wip_T(w1));
                arrW1 = step1 + step2 + step3;
                
                step4 = dis_matrix(wip_T(w1), wip_T(w2));
                arrW2 = arrW1 + step4;
                
                totDist   = step1 + step2 + step3 + step4;
                arrTimes  = [arrW1, arrW2];
                isFeasible = (arrW1 <= remaining_Q(w1)) && (arrW2 <= remaining_Q(w2));
                
            otherwise
                error('未知的策略類型: %s', strategy);
        end
    end
end

