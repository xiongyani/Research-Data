% �̶�·����Excel�ļ���
filename = 'C:\Users\Administrator\Desktop\ʵ������\ԭʼ����\S3.xlsx';

% ��ȡExcel�ļ��ĵ�16������ (P��)
data = readtable(filename);
y_values = data{:, 16};  % ��16������ (P��)

% ��ʼ���������
x_results = nan(size(y_values));

% ����ʽ���̵�ϵ��
coeffs = [-1005.2639, 622.0188, -136.2034, 12.5961, -0.0609];

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
    equation_coeffs(end) = equation_coeffs(end) - y;  % ���³�����Ϊ (-0.0609 - y)
    
    % ������ʽ����
    x_all = roots(equation_coeffs);
    
    % ɸѡ���ڷ�Χ[0, 0.15]�ڵ�ʵ��
    x_real = x_all(imag(x_all) == 0 & real(x_all) >= 0 & real(x_all) <= 0.15);
    
    % ���������Ч�⣬ȡ��һ���⣻���򣬱���ΪNaN
    if ~isempty(x_real)
        x_results(i) = x_real(1);
    end
end

% �� x_results ת��Ϊ����������ʹ�� xlswrite д�� Excel �ļ��� Q ��
xlswrite(filename, x_results(:), 'Sheet1', 'Q2'); % �� Q2 ��ʼд�룬�����һ��Ϊ��ͷ
disp('�����ѳɹ�д�� Excel �ļ��� Q ��');