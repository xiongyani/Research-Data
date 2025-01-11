% �̶�·����Excel�ļ���
filename = 'C:\Users\Administrator\Desktop\ʵ������\ԭʼ����\S2.xlsx';

% ��ȡExcel�ļ��ĵ�11������ (K��)
data = readtable(filename);
y_values = data{:, 11};  % ��11������ (K��)

% ��ʼ���������
x_results = nan(size(y_values));

% ����ʽ���̵�ϵ��
coeffs = [-970.7097, 550.2851, -112.0613, 9.7024, -0.063];

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
    equation_coeffs(end) = equation_coeffs(end) - y;  % ���³�����Ϊ (-0.063 - y)
    
    % ������ʽ����
    x_all = roots(equation_coeffs);
    
    % ɸѡ���ڷ�Χ[0, 0.15]�ڵ�ʵ��
    x_real = x_all(imag(x_all) == 0 & real(x_all) >= 0 & real(x_all) <= 0.15);
    
    % ���������Ч�⣬ȡ��һ���⣻���򣬱���ΪNaN
    if ~isempty(x_real)
        x_results(i) = x_real(1);
    end
end

% �� x_results ת��Ϊ����������ʹ�� xlswrite д�� Excel �ļ��� L ��
xlswrite(filename, x_results(:), 'Sheet1', 'L2'); % �� L2 ��ʼд�룬�����һ��Ϊ��ͷ
disp('�����ѳɹ�д�� Excel �ļ��� L ��');