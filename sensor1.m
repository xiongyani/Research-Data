% �̶�·����Excel�ļ���
filename = 'C:\Users\Administrator\Desktop\ʵ������\ԭʼ����\S1,2,3.xlsx';

% ��ȡExcel�ļ��ĵ�5������
data = readtable(filename);
y_values = data{:, 5};  % ��5������ (E��)

% ��ʼ���������
x_results = nan(size(y_values));

% ����ʽ���̵�ϵ��
coeffs = [-1233.3192, 474.3126, -68.5194, 4.586, -0.0237];

% ����ÿ��yֵ���������Ӧ��xֵ
for i = 1:length(y_values)
    y = y_values(i);
    
    % ��� y ֵ�Ƿ�Ϊ NaN �� Inf
    if isnan(y) || isinf(y)
        % ��� y ֵ��Ч�������ôμ���
        continue;
    end
    
    % �����̳������ȥyֵ���������з���Ϊ��׼��ʽ
    equation_coeffs = coeffs;
    equation_coeffs(end) = equation_coeffs(end) - y;  % ���³�����Ϊ (-0.0237 - y)
    
    % ������ʽ����
    x_all = roots(equation_coeffs);
    
    % ɸѡ���ڷ�Χ[0, 0.15]�ڵ�ʵ��
    x_real = x_all(imag(x_all) == 0 & real(x_all) >= 0 & real(x_all) <= 0.15);
    
    % ���������Ч�⣬ȡ��һ���⣻���򣬱���ΪNaN
    if ~isempty(x_real)
        x_results(i) = x_real(1);
    end
end

% �� x_results ת��Ϊ����������ʹ�� xlswrite д�� Excel �ļ��� F ��
xlswrite(filename, x_results(:), 'Sheet1', 'F2'); % �� F2 ��ʼд�룬�����һ��Ϊ��ͷ
disp('�����ѳɹ�д�� Excel �ļ��� F ��');