%%  Introduction program with CARD

% This program is used to realise automatic tube lens design with a pair of
% achromatic doublets from stock optics
%Lens catalog: name it as "All_Lenses_app.xlsx"
%pairs_number: (lens 1 number)0(lens 2 number) (eg. xxx0xxx, xx0xxx or x0xxx)
%file "BestRMSfield" includes the pairs_number and the RMS wavefront error in waves of the best pair of doublets.

% Copyright (C) 2021 Imperial College London.
% All rights reserved.
%
% This program is free software; you can redistribute it and/or modify
% it under the terms of the GNU General Public License as published by
% the Free Software Foundation; either version 2 of the License, or
% (at your option) any later version.
%
% This program is distributed in the hope that it will be useful,
% but WITHOUT ANY WARRANTY; without even the implied warranty of
% MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
% GNU General Public License for more details.
%
% You should have received a copy of the GNU General Public License along
% with this program; if not, write to the Free Software Foundation, Inc.,
% 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
%
% This software tool was developed with support from the UK 
% Engineering and Physical Sciences Council 
%
% Author: Wenzhi Hong


%% 

function [ r ] = tubelens_v4_3( args )

if ~exist('args', 'var')
    args = [];
end

% Initialize the OpticStudio connection
TheApplication = InitConnection();
if isempty(TheApplication)
    % failed to initialize a connection
    r = [];
else
    try
        r = BeginApplication(TheApplication, args);
        CleanupConnection(TheApplication);
    catch err
        CleanupConnection(TheApplication);
        rethrow(err);
    end
end
end


function [r] = BeginApplication(TheApplication, args)

import ZOSAPI.*;

    TheSystem = TheApplication.PrimarySystem;
    %set up primary optical system
    TheSystem = TheApplication.PrimarySystem;
    sampleDir = TheApplication.SamplesDir;
    
%% %% Diameters and focal lengths of doublets (input)
global cnt;%pair counter
cnt = 0;
global dis_b;%pair counter

n1 = 1;%Air
n2 = 1;%Air
n3 = 1;%Air
global best_pair;
global wave_max_a;
global wave_min_a;
%Tube lens focal length
global ft_a;
%Entrance pupil diameter
global epl_a;
%separation
global d_max_a;
global d_min_a;
%field angle
global fieldangle_a;
%file path
global file_path_a;
%file_path = '\test28\';
global pairs_number_a;
pairs_number = -1;
pairs_number_a = -1;
global done_a;
global telecentric_s;
 if (telecentric_s==1234)
     telecentric = 1;
 else
     telecentric = 0;
 end

global first_d;

global first_v;
firstv = first_v;

done_a = 'Loading';

global accuracy;    
global RMS; 
 if (RMS==2)
     rms = 2;
 else
     rms = 1;
 end
 
global Config; 


ac = accuracy;
ft = ft_a;
epl = epl_a;
d_max = d_max_a;
d_min = d_min_a;
fieldangle = fieldangle_a;
file_path = file_path_a;
first = first_d;

%%  Start
max_field = 8;
cutoff_field = -1;
poly_rms = 100;

 %set up primary optical system
            TheSystem = TheApplication.PrimarySystem;
            sampleDir = TheApplication.SamplesDir;
            
            %make a folder
            spFolder = [char(sampleDir) '\' file_path '\'];
            mkdir(spFolder);

Height = ft*(fieldangle/180)*pi;

%% Database of doublets

[catalogN, catalogT] = xlsread(['All_Lenses_app.xlsx']);

%% Choose doublets 
    
for i = 1:size(catalogN,1)
    
        lens_catalogN1(i,:) = catalogN(i,:);
        lens_catalogT1(i,1) = catalogT(i,1);
        
    
        lens_catalogN2(i,:) = catalogN(i,:);
        lens_catalogT2(i,1) = catalogT(i,1);
        
end
     
%% Step 1  Generate zmx file and Seidel coefficients.
p = 1;

for bz1 = 1:size(lens_catalogN1,1)
    
    if (lens_catalogN1(bz1,7)~= -1 && lens_catalogN1(bz1,6)~= 1 && wave_min_a>=lens_catalogN1(bz1,7) && wave_max_a<=lens_catalogN1(bz1,8))
        match1 = lens_catalogN1(bz1,10);
        

        
    for bz2 = 1:size(lens_catalogN2,1)
        
    if (lens_catalogN2(bz2,7)~= -1 && lens_catalogN2(bz2,6)~= 1 && wave_min_a>=lens_catalogN2(bz2,7) && wave_max_a<=lens_catalogN2(bz2,8))  


        match2 = lens_catalogN2(bz2,10);
 
    
    if((lens_catalogN1(bz1,3)>= 2*ft) && (lens_catalogN2(bz2,3)<=2*ft)||(lens_catalogN1(bz1,3)<=2*ft))
      

        n1 = 1;%Air
        n2 = 1;%Air
        n3 = 1;%Air
        f1 = lens_catalogN1(bz1,3);
        f2 = lens_catalogN2(bz2,3);
        d_test = (1/ft-(1/f1 +1/f2))/(-(1/f1)*(1/f2));
        delta1 = -n1/n2 *d_test*(1/f2)/(1/ft);
        l = ft-abs(delta1);
        u_bar1 = Height/ft;
        h1 = l*u_bar1;
        h2 = h1+epl/2;

        u_bar2 = u_bar1-h1*1/f1;
        h3 = u_bar2*d_test+h1;
        h4 = h3+epl/2;

        D_L1 = 2*h2;
        D_L2 = 2*h4;
        
        d1 = lens_catalogN1(bz1,4);
        d2 = lens_catalogN2(bz2,4);
      
      
      
       if((d_test<=d_max) && (d_test>=d_min))
       if((d1>=D_L1) && (d2>=D_L2))
            VendorName1 = cell2mat(lens_catalogT1(bz1,1));
            f1 = lens_catalogN1(bz1,3);
            a1 = lens_catalogN1(bz1,4);
            VendorName2 = cell2mat(lens_catalogT2(bz2,1));
            f2 = lens_catalogN2(bz2,3);
            a2 = lens_catalogN2(bz2,4);
            w1 = num2str(lens_catalogN1(bz1,1));    %file name
            w2 = num2str(lens_catalogN2(bz2,1));
            w = [w1,'_',num2str(match1),'_',w2,'_',num2str(match2)];
            W = str2double(w1)*1000000+match1*100000+str2double(w2)*10+match2;
    
            aperture_d = [a1, a2];
            effl = [f1, f2];
            entrance_pupil_diameter = epl;
            focal_length = ft;
    
    
            %set up primary optical system
            TheSystem = TheApplication.PrimarySystem;
            sampleDir = TheApplication.SamplesDir;
            

            testFile = System.String.Concat(sampleDir, ['\',file_path,'\',w,'.zmx']);
            TheSystem.New(false);
            TheSystem.SaveAs(testFile);
           
    
            %set Aperture
            TheSystemData = TheSystem.SystemData;
            TheSystemData.Aperture.ApertureValue = epl;
            %se Field
            Field_1 = TheSystemData.Fields.GetField(1);
            NewField2 = TheSystemData.Fields.AddField(0,fieldangle/2,1.0);
            NewField3 = TheSystemData.Fields.AddField(0,fieldangle,1.0);
            %set Wavelength
            slPreset = TheSystemData.Wavelengths.SelectWavelengthPreset(ZOSAPI.SystemData.WavelengthPreset.FdC_Visible);
    
    
        cnt = cnt +1;
    
       %Insert lens1
        LensCatalogs = TheSystem.Tools.OpenLensCatalogs();
        
        LensCatalogs.GetAllVendors();
        VendorName = VendorName1;
        LensCatalogs.SelectedVendor = VendorName;
        ElementNumber = 2;
        LensCatalogs.NumberOfElements = ElementNumber;
        LensCatalogs.UseEFL = true;
        LensCatalogs.MinEFL = f1; %set for spercific focal length
        LensCatalogs.MaxEFL = f1; 
        LensCatalogs.UseEPD = true;
        LensCatalogs.MinEPD = a1; %set for aperture
        LensCatalogs.MaxEPD = a1;
        LensCatalogs.IncShapeBi = false;
        LensCatalogs.IncShapeEqui = false;
        LensCatalogs.IncShapeMeniscus = false;
        LensCatalogs.IncShapePlano = false;
        LensCatalogs.IncTypeAspheric = false;
        LensCatalogs.IncTypeGRIN = false;
        LensCatalogs.IncTypeSpherical = true;
        LensCatalogs.Run();
        Matchlens1 = LensCatalogs.MatchingLenses;
           
        LensCatalogs.GetResult(match1-1).InsertLensSeq(2, true, true); 
                    
        
        %Insert lens2
        LensCatalogs.GetAllVendors();
        VendorName = VendorName2;
        LensCatalogs.SelectedVendor = VendorName;
        ElementNumber = 2;
        LensCatalogs.NumberOfElements = ElementNumber;
        LensCatalogs.UseEFL = true;
        LensCatalogs.MinEFL = f2; %set for spercific focal length
        LensCatalogs.MaxEFL = f2; 
        LensCatalogs.UseEPD = true;
        LensCatalogs.MinEPD = a2; %set for aperture
        LensCatalogs.MaxEPD = a2;
        LensCatalogs.IncShapeBi = true;
        LensCatalogs.IncShapeEqui = true;
        LensCatalogs.IncShapeMeniscus = true;
        LensCatalogs.IncShapePlano = true;
        LensCatalogs.IncTypeAspheric = true;
        LensCatalogs.IncTypeGRIN = true;
        LensCatalogs.IncTypeSpherical = true;
        LensCatalogs.IncTypeToroidal = true;

        LensCatalogs.Run();
        Matchlens2 = LensCatalogs.MatchingLenses;
        LensCatalogs.GetResult(match2-1).InsertLensSeq(5, true, true); 
     
        LensCatalogs.Close();
%% optimization
    if (Config==1)
        TheLDE = TheSystem.LDE;
        TheLDE.RunTool_ReverseElements(2,4);
        TheLDE.RunTool_ReverseElements(5,7);
    elseif (Config==2)
        TheLDE = TheSystem.LDE;
        TheLDE.RunTool_ReverseElements(2,4);
    elseif (Config==4)
        TheLDE = TheSystem.LDE;
        TheLDE.RunTool_ReverseElements(5,7);
    end
    
    if (telecentric==1)
        %telecentric
        
        %Get surfaces 
        TheLDE = TheSystem.LDE;
        Surface_1 = TheLDE.GetSurfaceAt(1);
        Surface_2 = TheLDE.GetSurfaceAt(2);
        Surface_3 = TheLDE.GetSurfaceAt(3);
        Surface_4 = TheLDE.GetSurfaceAt(4);
        Surface_5 = TheLDE.GetSurfaceAt(5);
        Surface_6 = TheLDE.GetSurfaceAt(6);
        Surface_7 = TheLDE.GetSurfaceAt(7);
        
       

        
%% %make thicknesses variable
        Surface_1.ThicknessCell.MakeSolveVariable();
        Surface_4.ThicknessCell.MakeSolveVariable();
        Surface_7.ThicknessCell.MakeSolveVariable();
        
        if rms == 2
            
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 2;%refer to chief ray
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();  
        
        else
        %Merit Functions
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 1;%refer to centroid
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();
        end
              
                
        Operand_2=TheMFE.InsertNewOperandAt(2);
        Operand_2.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.EFFL);
        Operand_2.Target = ft;
        Operand_2.Weight = 1.0;
        
        Operand_3=TheMFE.InsertNewOperandAt(3);
        Operand_3.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.CARD);
        Operand_3.Target = ft;
        Operand_3.Weight = 1.0;
        Operand_3.GetCellAt(6).IntegerValue = 4;

        %local optimisation till completion
        LocalOpt = TheSystem.Tools.OpenLocalOptimization();
        LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
        LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
        LocalOpt.NumberOfCores = 8;
        LocalOpt.RunAndWaitForCompletion();
        LocalOpt.Close();
    end
    
     if (firstv==2021 && telecentric ~= 1)
        % non-telecentric + target b value
        
        %Get surfaces 
        TheLDE = TheSystem.LDE;
        Surface_1 = TheLDE.GetSurfaceAt(1);
        Surface_2 = TheLDE.GetSurfaceAt(2);
        Surface_3 = TheLDE.GetSurfaceAt(3);
        Surface_4 = TheLDE.GetSurfaceAt(4);
        Surface_5 = TheLDE.GetSurfaceAt(5);
        Surface_6 = TheLDE.GetSurfaceAt(6);
        Surface_7 = TheLDE.GetSurfaceAt(7);
        
       

        
%% %make thicknesses variable
        Surface_1.ThicknessCell.MakeSolveVariable();
        Surface_4.ThicknessCell.MakeSolveVariable();
        Surface_7.ThicknessCell.MakeSolveVariable();

        if rms == 2
            
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 2;%refer to chief ray
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();  
        
        else
        %Merit Functions
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 1;%refer to centroid
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();
        end
        
        Operand_2=TheMFE.InsertNewOperandAt(2);
        Operand_2.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.EFFL);
        Operand_2.Target = ft;
        Operand_2.Weight = 1.0;
        
        Operand_3=TheMFE.InsertNewOperandAt(3);
        Operand_3.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.CARD);
        Operand_3.Target = dis_b;
        Operand_3.Weight = 1.0;
        Operand_3.GetCellAt(6).IntegerValue = 4;

        %local optimisation till completion
        LocalOpt = TheSystem.Tools.OpenLocalOptimization();
        LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
        LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
        LocalOpt.NumberOfCores = 8;
        LocalOpt.RunAndWaitForCompletion();
        LocalOpt.Close();
    end
    
    
    if (firstv==-9 && telecentric ~= 1)
        %fixed distance
        %Get surfaces
        TheLDE = TheSystem.LDE;
        Surface_1 = TheLDE.GetSurfaceAt(1);
        Surface_4 = TheLDE.GetSurfaceAt(4);
        Surface_7 = TheLDE.GetSurfaceAt(7);
        
        %set thickness of first surface
        Surface_1.Thickness = first;

        %make thicknesses variable
        
        Surface_4.ThicknessCell.MakeSolveVariable();
        Surface_7.ThicknessCell.MakeSolveVariable();

        if rms == 2
            
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 2;%refer to chief ray
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();  
        
        else
        %Merit Functions
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 1;%refer to centroid
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();
        end
        
        Operand_2=TheMFE.InsertNewOperandAt(2);
        Operand_2.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.EFFL);
        Operand_2.Target = ft;
        Operand_2.Weight = 1.0;
        

        %local optimisation till completion
        LocalOpt = TheSystem.Tools.OpenLocalOptimization();
        LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
        LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
        LocalOpt.NumberOfCores = 8;
        LocalOpt.RunAndWaitForCompletion();
        LocalOpt.Close();
    end

    if (firstv==-1 && telecentric~= 1)
        %variable
        %Get surfaces
        TheLDE = TheSystem.LDE;
        Surface_1 = TheLDE.GetSurfaceAt(1);
        Surface_4 = TheLDE.GetSurfaceAt(4);
        Surface_7 = TheLDE.GetSurfaceAt(7);

        %make thicknesses variable
        Surface_1.ThicknessCell.MakeSolveVariable();
        Surface_4.ThicknessCell.MakeSolveVariable();
        Surface_7.ThicknessCell.MakeSolveVariable();

        if rms == 2
            
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 2;%refer to chief ray
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();  
        
        else
        %Merit Functions
        TheMFE = TheSystem.MFE;
        OptWizard = TheMFE.SEQOptimizationWizard;
        OptWizard.Data = 0;%RMS WAVEFRONT
        OptWizard.OverallWeight = 1;
        OptWizard.Ring = 9;
        OptWizard.Reference = 1;%refer to centroid
        %Set air boundaries
        OptWizard.IsAirUsed = true;
        OptWizard.AirMin = 0.5;
        OptWizard.AirMax = 1000;
        OptWizard.AirEdge = 0.5;
        OptWizard.Apply();
        end
        
        Operand_2=TheMFE.InsertNewOperandAt(2);
        Operand_2.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.EFFL);
        Operand_2.Target = ft;
        Operand_2.Weight = 1.0;
        

        %local optimisation till completion
        LocalOpt = TheSystem.Tools.OpenLocalOptimization();
        LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
        LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
        LocalOpt.NumberOfCores = 8;
        LocalOpt.RunAndWaitForCompletion();
        LocalOpt.Close();
    end
    
    
    

        %Save and close
        TheSystem.Save();

%% Step 2 
    
    
    %open file
    testFile = System.String.Concat(sampleDir,['\',file_path,'\',w,'.zmx']);
    TheSystem.LoadFile(testFile,false);
    lens_no = 2;% number of lenses
    j = 2;
    lens_data = NaN(100,100);

    
    %% pull data from rms_field
    
if rms == 2
       % Set up primary optical system
    TheSystem = TheApplication.PrimarySystem;
    sampleDir = TheApplication.SamplesDir;
    %open file
    
    testFile = System.String.Concat(sampleDir, ['\',file_path,'\',w,'.zmx']);
    TheSystem.LoadFile(testFile,false);

    %create analysis
    TheAnalyses = TheSystem.Analyses;
    newWin = TheAnalyses.New_RMSField();
    
    %setting
    newWin_Settings = newWin.GetSettings();
    newWin_Settings.ShowDiffractionLimit = true;
    newWin_Settings.RayDensity = ZOSAPI.Analysis.Settings.RMS.RayDensities.RayDens_1024x1024;
    newWin_Settings.FieldDensity = ZOSAPI.Analysis.Settings.RMS.FieldDensities.FieldDens_100;
    newWin_Settings.ReferTo = ZOSAPI.Analysis.Settings.RMS.ReferTo.ChiefRay;
    
    %run analysis
    newWin.ApplyAndWaitForCompletion();
    newWin_Results = newWin.GetResults();
    

    %Read and plot data series
    dataSeries = newWin_Results.DataSeries;
    cc = lines(double(newWin_Results.NumberOfDataSeries));
    for gridN=1:newWin_Results.NumberOfDataSeries
        data = dataSeries(gridN);
        y = data.YData.Data.double;
        x = data.XData.Data.double;

    end
    

        cut_off(2,1) = y(size(y,1),1);
        cut_off(2,2) = y(size(y,1),2);
        cut_off(2,3) = y(size(y,1),3);
        cut_off(2,4) = y(size(y,1),4);

    
%% find cut-off field of poly curve

    %Add field
     TheSystemData = TheSystem.SystemData;
     NewField4 = TheSystemData.Fields.AddField(0,max_field,1.0);
    
    %create analysis
    TheAnalyses = TheSystem.Analyses;
    newWin = TheAnalyses.New_RMSField();
    
    %setting
    newWin_Settings = newWin.GetSettings();
    newWin_Settings.ShowDiffractionLimit = true;
    newWin_Settings.RayDensity = ZOSAPI.Analysis.Settings.RMS.RayDensities.RayDens_1024x1024;
    newWin_Settings.FieldDensity = ZOSAPI.Analysis.Settings.RMS.FieldDensities.FieldDens_100;
    newWin_Settings.ReferTo = ZOSAPI.Analysis.Settings.RMS.ReferTo.ChiefRay;
    
    %run analysis
    newWin.ApplyAndWaitForCompletion();
    newWin_Results = newWin.GetResults();
    
    
    %Read and plot data series
    dataSeries = newWin_Results.DataSeries;
    cc = lines(double(newWin_Results.NumberOfDataSeries));
    for gridN=1:newWin_Results.NumberOfDataSeries
        data = dataSeries(gridN);
        y = data.YData.Data.double;
        x = data.XData.Data.double;
    end
    


    cut_off(2,5) = 0;
    
    for i = 1:size(y,1)
        if (y(i,1) <= y(i,5))
            cut_off(2,5) = x(1,i);
        
        end
        
        
    end



if (cut_off(2,5)>=cutoff_field)
    cutoff_field = cut_off(2,5);
    if (cut_off(2,1)<poly_rms)
        poly_rms = cut_off(2,1);
        pairs_number(p,1) = W
        cut_off(2,6) = pairs_number(p,1);
        pairs_number_a = num2str(pairs_number);
        p = p+1;
        xlswrite(['BestRMSfield.xlsx'], cut_off, 1);%write data
        xlswrite(['Pairsnumber.xlsx'], pairs_number, 1);%write data

        txt = {'Poly(/wave)','0.4861mm(/wave)', '0.5876mm(/wave)','0.6563mm(/wave)','cut-off field andgle(/degree)','Pairs number'};
        xlswrite(['BestRMSfield.xlsx'],txt,1);
    end
    
end

else
    
       % Set up primary optical system
    TheSystem = TheApplication.PrimarySystem;
    sampleDir = TheApplication.SamplesDir;
    %open file
    
    testFile = System.String.Concat(sampleDir, ['\',file_path,'\',w,'.zmx']);
    TheSystem.LoadFile(testFile,false);

    %create analysis
    TheAnalyses = TheSystem.Analyses;
    newWin = TheAnalyses.New_RMSField();
    
    %setting
    newWin_Settings = newWin.GetSettings();
    newWin_Settings.ShowDiffractionLimit = true;
    newWin_Settings.RayDensity = ZOSAPI.Analysis.Settings.RMS.RayDensities.RayDens_1024x1024;
    newWin_Settings.FieldDensity = ZOSAPI.Analysis.Settings.RMS.FieldDensities.FieldDens_100;
    newWin_Settings.ReferTo = ZOSAPI.Analysis.Settings.RMS.ReferTo.Centroid;
    
    %run analysis
    newWin.ApplyAndWaitForCompletion();
    newWin_Results = newWin.GetResults();
    

    %Read and plot data series
    dataSeries = newWin_Results.DataSeries;
    cc = lines(double(newWin_Results.NumberOfDataSeries));
    for gridN=1:newWin_Results.NumberOfDataSeries
        data = dataSeries(gridN);
        y = data.YData.Data.double;
        x = data.XData.Data.double;

    end
    


        cut_off(2,1) = y(size(y,1),1);
        cut_off(2,2) = y(size(y,1),2);
        cut_off(2,3) = y(size(y,1),3);
        cut_off(2,4) = y(size(y,1),4);

    
%% find cut-off field of poly curve

    %Add field
     TheSystemData = TheSystem.SystemData;
     NewField4 = TheSystemData.Fields.AddField(0,max_field,1.0);
    
    %create analysis
    TheAnalyses = TheSystem.Analyses;
    newWin = TheAnalyses.New_RMSField();
    
    %setting
    newWin_Settings = newWin.GetSettings();
    newWin_Settings.ShowDiffractionLimit = true;
    newWin_Settings.RayDensity = ZOSAPI.Analysis.Settings.RMS.RayDensities.RayDens_1024x1024;
    newWin_Settings.FieldDensity = ZOSAPI.Analysis.Settings.RMS.FieldDensities.FieldDens_100;
    newWin_Settings.ReferTo = ZOSAPI.Analysis.Settings.RMS.ReferTo.Centroid;
    
    %run analysis
    newWin.ApplyAndWaitForCompletion();
    newWin_Results = newWin.GetResults();
    

    
    %Read and plot data series
    dataSeries = newWin_Results.DataSeries;
    cc = lines(double(newWin_Results.NumberOfDataSeries));
    for gridN=1:newWin_Results.NumberOfDataSeries
        data = dataSeries(gridN);
        y = data.YData.Data.double;
        x = data.XData.Data.double;

    end
    


    cut_off(2,5) = 0;
    
    for i = 1:size(y,1)
        if (y(i,1) <= y(i,5))
            cut_off(2,5) = x(1,i);
        
        end
        
        
    end



if (cut_off(2,5)>=cutoff_field)
    cutoff_field = cut_off(2,5);
    if (cut_off(2,1)<poly_rms)
        poly_rms = cut_off(2,1);
        pairs_number(p,1) = W
        cut_off(2,6) = pairs_number(p,1);
        pairs_number_a = num2str(pairs_number);
        p = p+1;
        xlswrite(['BestRMSfield.xlsx'], cut_off, 1);%write data
        xlswrite(['Pairsnumber.xlsx'], pairs_number, 1);%write data

        txt = {'Poly(/wave)','0.4861mm(/wave)', '0.5876mm(/wave)','0.6563mm(/wave)','cut-off field andgle(/degree)','Pairs number'};
        xlswrite(['BestRMSfield.xlsx'],txt,1);
    end
    
end
    

end

clear cut_off;


       end
    
    
       end
    end
    end
    end
   
    
    end
    
end


best_pair = pairs_number(size(pairs_number,1),1);    


  
    
    
  
    
r = [];

end

function app = InitConnection()

import System.Reflection.*;

% Find the installed version of OpticStudio.
zemaxData = winqueryreg('HKEY_CURRENT_USER', 'Software\Zemax', 'ZemaxRoot');
NetHelper = strcat(zemaxData, '\ZOS-API\Libraries\ZOSAPI_NetHelper.dll');
% Note -- uncomment the following line to use a custom NetHelper path
% NetHelper = 'C:\Users\AB\Documents\Zemax\ZOS-API\Libraries\ZOSAPI_NetHelper.dll';
% This is the path to OpticStudio
NET.addAssembly(NetHelper);

success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize();
% Note -- uncomment the following line to use a custom initialization path
% success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize('C:\Program Files\OpticStudio\');
if success == 1
    LogMessage(strcat('Found OpticStudio at: ', char(ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory())));
else
    app = [];
    return;
end

% Now load the ZOS-API assemblies
NET.addAssembly(AssemblyName('ZOSAPI_Interfaces'));
NET.addAssembly(AssemblyName('ZOSAPI'));

% Create the initial connection class
TheConnection = ZOSAPI.ZOSAPI_Connection();

% Attempt to create a Standalone connection

% NOTE - if this fails with a message like 'Unable to load one or more of
% the requested types', it is usually caused by try to connect to a 32-bit
% version of OpticStudio from a 64-bit version of MATLAB (or vice-versa).
% This is an issue with how MATLAB interfaces with .NET, and the only
% current workaround is to use 32- or 64-bit versions of both applications.
app = TheConnection.CreateNewApplication();
if isempty(app)
   HandleError('An unknown connection error occurred!');
end
if ~app.IsValidLicenseForAPI
    HandleError('License check failed!');
    app = [];
end

end

function LogMessage(msg)
disp(msg);
end

function HandleError(error)
ME = MException('zosapi:HandleError', error);
throw(ME);
end

function  CleanupConnection(TheApplication)
% Note - this will close down the connection.

% If you want to keep the application open, you should skip this step
% and store the instance somewhere instead.
TheApplication.CloseApplication();
end


