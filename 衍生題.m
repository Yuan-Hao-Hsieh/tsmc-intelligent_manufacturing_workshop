

%% ========== WIP排程問題解決方案（衍伸題） ==========
% 使用模擬退火算法解決TSMC智慧製造工作坊排程問題
% 衍伸題：每台車最多可負責4個WIP (0-4)
% 目標1：最小化違反等候時間限制的WIP數
% 目標2：最小化救貨車隊的移動距離

%% ========== 基本初始化 ==========
clc; clear; close all;
fprintf('WIP排程問題（衍伸題）解決方案\n');
fprintf('目標：最小化違反Q-time的WIP數量，同時最小化總移動距離\n\n');

% 設定檔案路徑 (依實際需要修改)
filePath = 'C:\Users\user\Desktop\113-2\TSMC WS\dataset_2B.xlsx';
% 設定參數
num_t_cart = 20; % 車輛數量
max_wip_per_cart = 4; % 每台車最多可負責的WIP數量（衍伸題）

%% ========== 讀取數據 ==========
fprintf('讀取數據中...\n');

% 讀取車輛資料
[~, ~, cart_raw] = xlsread(filePath, 'CART');
cart_raw = cart_raw(2:end, :); % 移除標題行
cart_ids = cart_raw(:, 1);
ini_cart_loc = cell2mat(cart_raw(:, 2));

% 讀取距離矩陣
dis_matrix = readmatrix(filePath, 'Sheet', 'Distance_Matrix', 'Range', 'B2:AY51');

% 讀取紅色群組資料 (優先度高的WIP)
red_remaining_Q = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'K45:K57');
red_wip_F = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'L45:L57');
red_wip_T = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'M45:M57');
red_wip_ids = readcell(filePath, 'Sheet', 'WIP', 'Range', 'J45:J57');
num_red_wip = length(red_wip_ids);

% 讀取淺綠群組資料
green_remaining_Q = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'K59:K72');
green_wip_F = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'L59:L72');
green_wip_T = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'M59:M72');
green_wip_ids = readcell(filePath, 'Sheet', 'WIP', 'Range', 'J59:J72');
num_green_wip = length(green_wip_ids);

% 讀取深綠群組資料
dgreen_remaining_Q = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'K74:K86');
dgreen_wip_F = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'L74:L86');
dgreen_wip_T = readmatrix(filePath, 'Sheet', 'WIP', 'Range', 'M74:M86');
dgreen_wip_ids = readcell(filePath, 'Sheet', 'WIP', 'Range', 'J74:J86');
num_dgreen_wip = length(dgreen_wip_ids);

% 合併所有WIP數據
all_wip_ids = [red_wip_ids; green_wip_ids; dgreen_wip_ids];
all_remaining_Q = [red_remaining_Q; green_remaining_Q; dgreen_remaining_Q];
all_wip_F = [red_wip_F; green_wip_F; dgreen_wip_F];
all_wip_T = [red_wip_T; green_wip_T; dgreen_wip_T];
num_total_wip = num_red_wip + num_green_wip + num_dgreen_wip;

fprintf('成功讀取數據：%d台推車, %d個WIP\n', num_t_cart, num_total_wip);
fprintf('  - 紅色(高優先級)WIP: %d個\n', num_red_wip);
fprintf('  - 淺綠WIP: %d個\n', num_green_wip);
fprintf('  - 深綠WIP: %d個\n', num_dgreen_wip);

%% ========== 初始解生成 ==========
fprintf('\n生成初始解...\n');

% 創建WIP索引陣列 (1-40)
wip_indices = 1:num_total_wip;

% 根據Q-Time對WIP進行排序 (從小到大，優先考慮時間緊迫的WIP)
[sorted_Q_time, sorted_idx] = sort(all_remaining_Q);

% 初始化車輛分配
cart_assignment = cell(num_t_cart, 1);
for i = 1:num_t_cart
    cart_assignment{i} = [];
end

% 貪婪分配策略：優先分配Q-time小的WIP
wip_index = 1;
for cart_idx = 1:num_t_cart
    % 每台車隨機分配1-4個WIP (衍伸題)
    if wip_index <= num_total_wip
        num_assigned = min(randi([1, max_wip_per_cart]), num_total_wip - wip_index + 1);
        for j = 1:num_assigned
            if wip_index <= num_total_wip
                cart_assignment{cart_idx} = [cart_assignment{cart_idx}, sorted_idx(wip_index)];
                wip_index = wip_index + 1;
            end
        end
    end
end

% 確保所有WIP都被分配（檢查是否有遺漏）
assigned_wips = [];
for i = 1:num_t_cart
    assigned_wips = [assigned_wips, cart_assignment{i}];
end
assigned_wips = sort(assigned_wips);

if length(assigned_wips) < num_total_wip
    fprintf('警告：初始解中只分配了 %d/%d 個WIP！修正中...\n', length(assigned_wips), num_total_wip);
    
    % 找出未分配的WIP
    unassigned_wips = setdiff(1:num_total_wip, assigned_wips);
    
    % 將未分配的WIP分配到負載最輕的車輛
    for wip_idx = unassigned_wips
        % 尋找負載最輕的車輛
        min_load = max_wip_per_cart;
        min_cart_idx = 1;
        for cart_idx = 1:num_t_cart
            if length(cart_assignment{cart_idx}) < min_load
                min_load = length(cart_assignment{cart_idx});
                min_cart_idx = cart_idx;
            end
        end
        
        % 分配WIP
        cart_assignment{min_cart_idx} = [cart_assignment{min_cart_idx}, wip_idx];
    end
    
    % 再次確認
    assigned_wips = [];
    for i = 1:num_t_cart
        assigned_wips = [assigned_wips, cart_assignment{i}];
    end
    fprintf('修正後：已分配 %d/%d 個WIP\n', length(assigned_wips), num_total_wip);
end

% 驗證每台車的WIP數量不超過限制
for i = 1:num_t_cart
    if length(cart_assignment{i}) > max_wip_per_cart
        fprintf('警告：車輛 %d 分配了 %d 個WIP，超過了 %d 個的限制！\n', ...
            i, length(cart_assignment{i}), max_wip_per_cart);
    end
end

%% ========== 模擬退火算法參數設定 ==========
% 模擬退火算法參數
initial_temp = 2000;     % 初始溫度
cooling_rate = 0.8;     % 冷卻率
min_temp = 1;            % 最小溫度
max_iterations = 200;    % 每個溫度的最大迭代次數

%% ========== 模擬退火算法主函數 ==========
fprintf('\n開始執行模擬退火算法...\n');

% 評估初始解
[current_violations, current_distance, violation_details] = evaluate_solution(cart_assignment, ini_cart_loc, all_wip_F, all_wip_T, all_remaining_Q, dis_matrix, all_wip_ids);

% 初始化最佳解
best_assignment = cart_assignment;
best_violations = current_violations;
best_distance = current_distance;
best_violation_details = violation_details;

% 記錄結果
result_history = [0, best_violations, best_distance];

fprintf('初始解: 違反Q-Time的WIP數量 = %d, 總距離 = %.2f\n', best_violations, best_distance);
if best_violations > 0
    fprintf('違反Q-Time的WIPs: ');
    for i = 1:length(best_violation_details)
        fprintf('%s ', best_violation_details{i});
    end
    fprintf('\n');
end

% 主迭代循環
temp = initial_temp;
iteration = 0;
no_improvement_count = 0;
max_no_improvement = 10;  % 連續10次無改進則提前結束

while temp > min_temp && no_improvement_count < max_no_improvement
    improvement_in_temp = false;
    
    for i = 1:max_iterations
        % 生成新的解
        new_assignment = generate_neighbor(cart_assignment, num_total_wip, max_wip_per_cart);
        
        % 評估新解
        [new_violations, new_distance, new_violation_details] = evaluate_solution(new_assignment, ini_cart_loc, all_wip_F, all_wip_T, all_remaining_Q, dis_matrix, all_wip_ids);
        
        % 計算目標函數差值 (主要優化違反數量，次要優化距離)
        if new_violations < current_violations
            accept = true;
        elseif new_violations == current_violations && new_distance < current_distance
            accept = true;
        else
            % 計算接受概率 (優先考慮違反數量)
            delta = (new_violations - current_violations) * 1000 + (new_distance - current_distance) / 100;
            accept_prob = exp(-delta / temp);
            accept = rand() < accept_prob;
        end
        
        % 接受或拒絕新解
        if accept
            cart_assignment = new_assignment;
            current_violations = new_violations;
            current_distance = new_distance;
            
            % 更新最佳解
            if new_violations < best_violations || (new_violations == best_violations && new_distance < best_distance)
                best_assignment = new_assignment;
                best_violations = new_violations;
                best_distance = new_distance;
                best_violation_details = new_violation_details;
                improvement_in_temp = true;
            end
        end
    end
    
    % 記錄結果
    iteration = iteration + 1;
    result_history = [result_history; iteration, best_violations, best_distance];
    
    % 更新無改進計數
    if improvement_in_temp
        no_improvement_count = 0;
    else
        no_improvement_count = no_improvement_count + 1;
    end
    
    % 降溫並顯示進度
    temp = temp * cooling_rate;
    fprintf('迭代 %d: 溫度 = %.2f, 違反數量 = %d, 總距離 = %.2f\n', iteration, temp, best_violations, best_distance);
    
    if no_improvement_count > 0
        fprintf('  (連續 %d 次迭代無改進)\n', no_improvement_count);
    end
end

fprintf('\n最終結果: 違反Q-Time的WIP數量 = %d, 總距離 = %.2f\n', best_violations, best_distance);
if best_violations > 0
    fprintf('違反Q-Time的WIPs: ');
    for i = 1:length(best_violation_details)
        fprintf('%s ', best_violation_details{i});
    end
    fprintf('\n');
end

%% ========== 生成排程計劃 ==========
fprintf('\n生成最終排程計劃...\n');
[final_schedule, detailed_schedule] = generate_schedule(best_assignment, ini_cart_loc, all_wip_F, all_wip_T, all_wip_ids, all_remaining_Q, dis_matrix);

% 確認排程不違反Q-time限制
fprintf('\n檢查排程是否違反Q-time限制...\n');
violation_count = 0;
for i = 1:size(detailed_schedule, 1)
    if strcmp(detailed_schedule{i, 5}, 'DELIVERY')
        wip_idx = detailed_schedule{i, 7};
        complete_time = detailed_schedule{i, 6};
        q_time = all_remaining_Q(wip_idx);
        
        if complete_time > q_time
            violation_count = violation_count + 1;
            fprintf('警告：WIP %s 的完成時間 %.2f 超過了Q-time %.2f\n', ...
                all_wip_ids{wip_idx}, complete_time, q_time);
        end
    end
end

if violation_count == 0
    fprintf('排程檢查通過：所有WIP都在Q-time限制內完成！\n');
else
    fprintf('排程檢查失敗：有 %d 個WIP違反了Q-time限制。\n', violation_count);
end

% 檢查是否所有WIP都被處理
scheduled_wips = unique(final_schedule(:, 3));
if length(scheduled_wips) < num_total_wip
    fprintf('警告：排程中只包含了 %d/%d 個WIP！\n', length(scheduled_wips), num_total_wip);
    
    % 檢查哪些WIP缺失
    all_wip_ids_cell = all_wip_ids;
    missing_wips = setdiff(all_wip_ids_cell, scheduled_wips);
    fprintf('缺失的WIPs: ');
    for i = 1:length(missing_wips)
        fprintf('%s ', missing_wips{i});
    end
    fprintf('\n');
else
    fprintf('所有 %d 個WIP都已納入排程。\n', num_total_wip);
end

% 轉換為表格
cart_id_col = final_schedule(:, 1);
order_col = cell2mat(final_schedule(:, 2));
wip_id_col = final_schedule(:, 3);
action_col = final_schedule(:, 4);
complete_time_col = cell2mat(final_schedule(:, 5));

schedule_table = table(cart_id_col, order_col, wip_id_col, action_col, complete_time_col, ...
    'VariableNames', {'CART_ID', 'ORDER', 'WIP_ID', 'ACTION', 'COMPLETE_TIME'});

% 排序 (按照CART_ID和ORDER)
schedule_table = sortrows(schedule_table, {'CART_ID', 'ORDER'});

% 寫入Excel文件
writetable(schedule_table, 'WIP_Schedule_Solution.xlsx');

fprintf('排程計劃已生成並儲存至 WIP_Schedule_Solution.xlsx\n');

%% ========== 可視化結果 ==========
figure;
subplot(2, 1, 1);
plot(result_history(:, 1), result_history(:, 2), 'r-o', 'LineWidth', 2);
title('迭代過程中違反Q-Time的WIP數量變化');
xlabel('迭代次數');
ylabel('違反數量');
grid on;

subplot(2, 1, 2);
plot(result_history(:, 1), result_history(:, 3), 'b-o', 'LineWidth', 2);
title('迭代過程中總距離變化');
xlabel('迭代次數');
ylabel('總距離');
grid on;

%% ========== 輔助函數 ==========
function [violations, total_distance, violation_details] = evaluate_solution(assignment, cart_locations, wip_from, wip_to, q_times, distance_matrix, wip_ids)
    % 評估解決方案，計算違反Q-Time的WIP數量和總距離
    violations = 0;
    total_distance = 0;
    violation_details = {};
    
    for cart_idx = 1:length(assignment)
        wips = assignment{cart_idx};
        if isempty(wips)
            continue;  % 跳過空車
        end
        
        current_time = 0;
        current_loc = cart_locations(cart_idx);
        
        for i = 1:length(wips)
            wip_idx = wips(i);
            from_loc = wip_from(wip_idx);
            to_loc = wip_to(wip_idx);
            q_time = q_times(wip_idx);
            
            % 從當前位置到WIP起點的時間
            pickup_time = distance_matrix(current_loc, from_loc);
            current_time = current_time + pickup_time;
            
            % 從WIP起點到終點的時間
            delivery_time = distance_matrix(from_loc, to_loc);
            current_time = current_time + delivery_time;
            
            % 檢查是否違反Q-Time
            if current_time > q_time
                violations = violations + 1;
                if ~isempty(wip_ids)
                    violation_details{end+1} = wip_ids{wip_idx};
                end
            end
            
            % 計算總距離
            total_distance = total_distance + pickup_time + delivery_time;
            
            % 更新當前位置
            current_loc = to_loc;
        end
    end
end

function new_assignment = generate_neighbor(assignment, num_wips, max_wips_per_cart)
    % 生成鄰近解
    new_assignment = assignment;
    
    % 隨機選擇修改類型
    mod_type = randi(4);
    
    switch mod_type
        case 1  % 兩台車之間交換WIP
            cart1 = randi(length(assignment));
            cart2 = randi(length(assignment));
            
            % 確保兩台車都有WIP
            if ~isempty(new_assignment{cart1}) && ~isempty(new_assignment{cart2})
                wip1_idx = randi(length(new_assignment{cart1}));
                wip2_idx = randi(length(new_assignment{cart2}));
                
                % 交換WIP
                temp = new_assignment{cart1}(wip1_idx);
                new_assignment{cart1}(wip1_idx) = new_assignment{cart2}(wip2_idx);
                new_assignment{cart2}(wip2_idx) = temp;
            end
            
        case 2  % 從一台車移動WIP到另一台車
            from_cart = randi(length(assignment));
            to_cart = randi(length(assignment));
            
            % 確保源車有WIP且目標車未滿載
            if ~isempty(new_assignment{from_cart}) && length(new_assignment{to_cart}) < max_wips_per_cart
                wip_idx = randi(length(new_assignment{from_cart}));
                
                % 移動WIP
                wip_to_move = new_assignment{from_cart}(wip_idx);
                new_assignment{from_cart}(wip_idx) = [];
                new_assignment{to_cart} = [new_assignment{to_cart}, wip_to_move];
            end
            
        case 3  % 變更車內WIP的順序
            cart = randi(length(assignment));
            
            % 確保車內有至少2個WIP
            if length(new_assignment{cart}) >= 2
                % 打亂順序
                new_assignment{cart} = new_assignment{cart}(randperm(length(new_assignment{cart})));
            end
            
        case 4  % 添加或移除WIP (衍伸題特有)
            cart = randi(length(assignment));
            
            if rand() < 0.5 && length(new_assignment{cart}) < max_wips_per_cart
                % 添加一個尚未分配的WIP
                all_assigned_wips = [];
                for i = 1:length(new_assignment)
                    all_assigned_wips = [all_assigned_wips, new_assignment{i}];
                end
                
                unassigned_wips = setdiff(1:num_wips, all_assigned_wips);
                
                if ~isempty(unassigned_wips)
                    wip_to_add = unassigned_wips(randi(length(unassigned_wips)));
                    new_assignment{cart} = [new_assignment{cart}, wip_to_add];
                end
            elseif ~isempty(new_assignment{cart})
                % 移除一個WIP
                wip_idx = randi(length(new_assignment{cart}));
                wip_to_remove = new_assignment{cart}(wip_idx);
                new_assignment{cart}(wip_idx) = [];
                
                % 確保被移除的WIP會被重新分配
                % 找出負載最輕的車輛
                min_load = max_wips_per_cart;
                min_cart_idx = 1;
                for i = 1:length(new_assignment)
                    if i ~= cart && length(new_assignment{i}) < min_load
                        min_load = length(new_assignment{i});
                        min_cart_idx = i;
                    end
                end
                
                % 如果有可用車輛，將WIP分配給它
                if min_load < max_wips_per_cart
                    new_assignment{min_cart_idx} = [new_assignment{min_cart_idx}, wip_to_remove];
                else
                    % 如果所有車輛都已滿載，隨機選擇一個不同的車輛
                    available_carts = setdiff(1:length(new_assignment), cart);
                    if ~isempty(available_carts)
                        random_cart = available_carts(randi(length(available_carts)));
                        new_assignment{random_cart} = [new_assignment{random_cart}, wip_to_remove];
                    else
                        % 如果沒有其他車輛，放回原車
                        new_assignment{cart} = [new_assignment{cart}, wip_to_remove];
                    end
                end
            end
    end
    
    % 確保沒有WIP丟失
    all_wips = [];
    for i = 1:length(new_assignment)
        all_wips = [all_wips, new_assignment{i}];
    end
    all_wips = sort(all_wips);
    
    if length(all_wips) ~= num_wips
        % 找出丟失的WIP
        missing_wips = setdiff(1:num_wips, all_wips);
        
        % 將丟失的WIP重新分配
        for wip_idx = missing_wips
            % 找出負載最輕的車輛
            min_load = max_wips_per_cart;
            min_cart_idx = 1;
            for i = 1:length(new_assignment)
                if length(new_assignment{i}) < min_load
                    min_load = length(new_assignment{i});
                    min_cart_idx = i;
                end
            end
            
            % 分配WIP
            new_assignment{min_cart_idx} = [new_assignment{min_cart_idx}, wip_idx];
        end
    end
end

function [schedule, detailed_schedule] = generate_schedule(assignment, cart_locations, wip_from, wip_to, wip_ids, q_times, distance_matrix)
    % 生成排程計劃，並返回詳細排程信息
    schedule = {};
    detailed_schedule = {};  % [Cart_ID, Order, WIP_ID, Action, Complete_Time, WIP_Index, Q-Time]
    
    for cart_idx = 1:length(assignment)
        wips = assignment{cart_idx};
        if isempty(wips)
            continue;  % 跳過空車
        end
        
        cart_id = sprintf('C%02d', cart_idx);
        current_time = 0;
        current_loc = cart_locations(cart_idx);
        
        for i = 1:length(wips)
            wip_idx = wips(i);
            wip_id = wip_ids{wip_idx};
            from_loc = wip_from(wip_idx);
            to_loc = wip_to(wip_idx);
            q_time = q_times(wip_idx);
            
            % 從當前位置到WIP起點的時間
            pickup_time = distance_matrix(current_loc, from_loc);
            current_time = current_time + pickup_time;
            
            % PICKUP 動作
            schedule = [schedule; {cart_id, i*2-1, wip_id, 'PICKUP', current_time}];
            detailed_schedule = [detailed_schedule; {cart_id, i*2-1, wip_id, 'PICKUP', 'PICKUP', current_time, wip_idx, q_time}];
            
            % 更新當前位置
            current_loc = from_loc;
            
            % 從WIP起點到終點的時間
            delivery_time = distance_matrix(from_loc, to_loc);
            current_time = current_time + delivery_time;
            
            % DELIVERY 動作
            schedule = [schedule; {cart_id, i*2, wip_id, 'DELIVERY', current_time}];
            detailed_schedule = [detailed_schedule; {cart_id, i*2, wip_id, 'DELIVERY', 'DELIVERY', current_time, wip_idx, q_time}];
            
            % 更新當前位置
            current_loc = to_loc;
        end
    end
end