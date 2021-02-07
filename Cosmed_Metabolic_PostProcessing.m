% This Matlab script calculates and analyzes the metabolic data from Cosmed. 
% The code uses Brockway equation (1987).
% Developer: Saman Madinei (2019) - Occupational Biomechanics and Ergonomics Labs at Virginia Tech


set(0,'defaultaxesfontsize',14);
set(0,'defaulttextfontsize',14);

set(0,'defaultaxeslinewidth',2);
set(0,'defaultlinelinewidth',2);

set(0,'defaultAxesFontName', 'Arial')
set(0,'defaultTextFontName', 'Arial')

files1 = {
    'LUPE1', 'Subject 1 - 0%', 'without_first', 1; ...
    'LUPE2', 'Subject 1 - 20%', 'without_first', 2; ...
    'JACK1', 'Subject 2 - 0%', 'without_first', 3; ...
    'JACK2', 'Subject 2 - 20%', 'without_first', 4};
files2 = {
    'DANIEL2', 'Subject 3 - 0%', 'with_first', 1; ...
    'DANIEL1', 'Subject 3 - 20%', 'with_first', 2; ...
    'ZIAD1', 'Subject 4 - 0%', 'with_first', 3; ...
    'ZIAD2', 'Subject 4 - 20%', 'with_first', 4};

files = files2;

for i = 1:size(files,1)
    [a,b] = xlsread(files{i,1});
    
    % Select cell range for numerical data
    data = a(4:end,10:end);
    
    %% Load time strings and convert to seconds
    time_string = b(4:end,10);
    [~,~,~,hours, minutes, sec] = datevec(time_string);
    time_sec = hours*3600 + minutes*60 + sec;
    
    [data_resampled,time_resampled] = resample(data(:,4:5),time_sec); %resample data for vo2 and vco2
    
    fc = .5;
    fs = 3;
    [b2,a2] = butter(4,fc/(fs/2));
    data_filtered = filtfilt(b2,a2,data_resampled);
    
    
    %% Get V02, VC02, convert ml/min to ml/s
    
    vo2 = 1/60 * data_filtered(:,1);
    vco2 = 1/60 * data_filtered(:,2);
    
    % T0_line = input('Starting line ');
    % %T0_s = input ('Starting time (sec)');
    % Tf_line = input('Ending line ');
    % %Tf_s = input ('Ending time (sec)');
    
    T0_s_1 = 180; %'00:03:00';
    Tf_s_1 = 300; %'00:05:00';
    T0_s_2 = 390; %'00:06:30';
    Tf_s_2 = 510; %'00:08:30';
    T0_s_3 = 600; %'00:10:00';
    Tf_s_3 = 720; %'00:12:00';
    
    % [~,T0_line_1] = min(abs(datenum(T0_s_1) - datenum(time_vec)));
    % [~,Tf_line_1] = min(abs(datenum(Tf_s_1) - datenum(time_vec)));
    % [~,T0_line_2] = min(abs(datenum(T0_s_2) - datenum(time_vec)));
    % [~,Tf_line_2] = min(abs(datenum(Tf_s_2) - datenum(time_vec)));
    % [~,T0_line_3] = min(abs(datenum(T0_s_3) - datenum(time_vec)));
    % [~,Tf_line_3] = min(abs(datenum(Tf_s_3) - datenum(time_vec)));
    
    [~,T0_line_1] = min(abs(T0_s_1 - time_resampled));
    [~,Tf_line_1] = min(abs(Tf_s_1 - time_resampled));
    [~,T0_line_2] = min(abs(T0_s_2 - time_resampled));
    [~,Tf_line_2] = min(abs(Tf_s_2 - time_resampled));
    [~,T0_line_3] = min(abs(T0_s_3 - time_resampled));
    [~,Tf_line_3] = min(abs(Tf_s_3 - time_resampled));
    
    % metabolic_power = 16.5*vo2+4.62*vco2; % equation from Weir 1949 with Nitrogen term removed
    metabolic_power = 16.58*vo2+4.51*vco2; % equation from Brockway 1987 with Nitro term removed 
    % reference: Brockway, 1987 Hum Nutr Clin Nutr 1987, 41: 463–471.
    
    BW = a(7,1);
    
    files{i,2}
    Avg_power1 = mean (metabolic_power(T0_line_1:Tf_line_1,1));
    COT1 = mean (data(T0_line_1:Tf_line_1,10));
    % BW = input ('Insert BW ');
    Avg_power_kg1 = Avg_power1/BW
    
    Avg_power2 = mean (metabolic_power(T0_line_2:Tf_line_2,1));
    COT2 = mean (data(T0_line_2:Tf_line_2,10));
    % BW = input ('Insert BW ');
    Avg_power_kg2 = Avg_power2/BW
    
    Avg_power3 = mean (metabolic_power(T0_line_3:Tf_line_3,1));
    COT3 = mean (data(T0_line_3:Tf_line_3,10));
    % BW = input ('Insert BW ');
    Avg_power_kg3 = Avg_power3/BW
    
    
    
    
    
    if strcmp(files{i,3}, 'without_first')
        figure(1)
        cmap = get(gca,'ColorOrder');
        color_1 = cmap(2,:);
        color_2 = cmap(4,:);
        disp_name1 = 'Without Exo';
        disp_name2 = 'With Exo';
        
    else
        figure(2)
        cmap = get(gca,'ColorOrder');
        color_1 = cmap(4,:);
        color_2 = cmap(2,:);
        disp_name1 = 'With Exo';
        disp_name2 = 'Without Exo';
        
    end
    
    subplot(2,2,files{i,4})
    h(1) = plot(time_resampled, metabolic_power, 'Color', cmap(1,:));
    hold on
    
    % plot(time_resampled, 16.58*(1/60*data_filtered(:,1))+4.51*(1/60*data_resampled(:,2)), 'Color', cmap(2,:))
    h(2) = plot([T0_s_1, Tf_s_1], [Avg_power1 Avg_power1], 'Color', color_1, 'DisplayName', disp_name1);
    h(3) = plot([T0_s_2, Tf_s_2], [Avg_power2 Avg_power2], 'Color', color_2, 'DisplayName', disp_name2);
    plot([T0_s_3, Tf_s_3], [Avg_power3 Avg_power3], 'Color', color_1)
    
    plot([T0_s_1, T0_s_1],[0,1000],'Color', color_1,'LineStyle', '--')
    plot([Tf_s_1, Tf_s_1],[0,1000],'Color', color_1,'LineStyle', '--')
    plot([T0_s_2, T0_s_2],[0,1000],'Color', color_2,'LineStyle', '--')
    plot([Tf_s_2, Tf_s_2],[0,1000],'Color', color_2,'LineStyle', '--')
    plot([T0_s_3, T0_s_3],[0,1000],'Color', color_1,'LineStyle', '--')
    plot([Tf_s_3, Tf_s_3],[0,1000],'Color', color_1,'LineStyle', '--')
    
    ylim([min(metabolic_power),max(metabolic_power)])
    xlim([0 750])
    title(files{i,2}, 'FontSize', 18)
    xlabel('Time (sec)', 'FontSize', 15)
    ylabel('Metabolic Power (Watts)', 'FontSize', 15)
    
    
    if files{i,4} == 2
        legend(h(2:3));
    end
    
end
