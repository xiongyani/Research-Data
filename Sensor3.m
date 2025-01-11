% 固定路径的Excel文件名
filename = 'C:\Users\Administrator\Desktop\实地试验\原始数据\S3.xlsx';

% 读取Excel文件的第16列数据 (P列)
data = readtable(filename);
y_values = data{:, 16};  % 第16列数据 (P列)

% 初始化结果数组
x_results = nan(size(y_values));

% 多项式方程的系数
coeffs = [-1005.2639, 622.0188, -136.2034, 12.5961, -0.0609];

% 遍历每个y值，逐个求解对应的x值
for i = 1:length(y_values)
    y = y_values(i);
    
    % 检查 y 值是否为 NaN 或 Inf
    if isnan(y) || isinf(y)
        % 如果 y 值无效，跳过该次计算
        continue;
    end
    
    % 将方程常数项减去y值，重新排列方程为标准形式
    equation_coeffs = coeffs;
    equation_coeffs(end) = equation_coeffs(end) - y;  % 更新常数项为 (-0.0609 - y)
    
    % 求解多项式方程
    x_all = roots(equation_coeffs);
    
    % 筛选出在范围[0, 0.15]内的实根
    x_real = x_all(imag(x_all) == 0 & real(x_all) >= 0 & real(x_all) <= 0.15);
    
    % 如果存在有效解，取第一个解；否则，保持为NaN
    if ~isempty(x_real)
        x_results(i) = x_real(1);
    end
end

% 将 x_results 转换为列向量，并使用 xlswrite 写入 Excel 文件的 Q 列
xlswrite(filename, x_results(:), 'Sheet1', 'Q2'); % 从 Q2 开始写入，假设第一行为表头
disp('数据已成功写入 Excel 文件的 Q 列');