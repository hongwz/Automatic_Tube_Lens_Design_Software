%% Lens data readout program

% This program is used to read out the lens details in Zemax Libraries and
% set flag values to each lens.

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
function [ r ] = Readingout_lenscatalogs( args )

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
    
%% read out all lens
%set up primary optical system
    TheSystem = TheApplication.PrimarySystem;
    sampleDir = TheApplication.SamplesDir;

    global focal_min;
    global focal_max;
    global a_min;
    global a_max;
    global flag;
    global vendorname;


    
    flag = 0;
    o1 = 0;

    
    Lens_catalog = cell(999,10);

  
    
    bz1 = 1;
    k = 0;
   

    
    
    
   for i = 1:size(vendorname,2)
    
    %make a file
    testFile = System.String.Concat(sampleDir, ['\Sequential\Objectives\vendorname_',cell2mat(vendorname(i)),'.zmx']);     
    TheSystem.New(false);
    TheSystem.SaveAs(testFile);
    
            

            
            
        %open lens catalogs and setting
    
            LensCatalogs = TheSystem.Tools.OpenLensCatalogs();
            LensCatalogs.GetAllVendors();
            VendorName = cell2mat(vendorname(i));
            LensCatalogs.SelectedVendor = VendorName;
            ElementNumber = 2;
            LensCatalogs.NumberOfElements = ElementNumber;
            LensCatalogs.UseEFL = true;
            LensCatalogs.MinEFL = focal_min; %set for spercific focal length
            LensCatalogs.MaxEFL = focal_max; 
            LensCatalogs.UseEPD = true;
            LensCatalogs.MinEPD = a_min; %set for aperture
            LensCatalogs.MaxEPD = a_max;
            LensCatalogs.IncShapeBi = false;
            LensCatalogs.IncShapeEqui = false;
            LensCatalogs.IncShapeMeniscus = false;
            LensCatalogs.IncShapePlano = false;
            LensCatalogs.IncTypeAspheric = false;
            LensCatalogs.IncTypeGRIN = false;
            LensCatalogs.IncTypeSpherical = true;
            LensCatalogs.IncTypeToroidal = false;

            LensCatalogs.Run();
            Matchlens = LensCatalogs.MatchingLenses;


            if (Matchlens>=1)
                
                for lens = 1:Matchlens
                    if(strcmp(VendorName,'BEFORT-WETZLAR'))
                        w1 = 0.400;
                        w2 = 0.700;
                        Lens_catalog(bz1,7) = num2cell(w1);
                        Lens_catalog(bz1,8) = num2cell(w2);
                    end
                    
                    if(strcmp(VendorName,'COMAR'))

                        w1 = 0.400;
                        w2 = 0.700;
                        Lens_catalog(bz1,7) = num2cell(w1);
                        Lens_catalog(bz1,8) = num2cell(w2);
                    end
                    
                    if(strcmp(VendorName,'CVI MELLES GRIOT'))
                        name = char(LensCatalogs.GetResult(lens-1).LensName); 
                        if(strfind(name,'LAI'))
                            w1 = 0.707;
                            w2 = 1.015;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strfind(name,'AAP'))
                            w1 = 0.400;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strfind(name,'LAO'))
                            w1 = 0.400;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        else
                            Lens_catalog(bz1,7) = num2cell(-1);
                            Lens_catalog(bz1,8) = num2cell(-1);
                        end
                        
                    end
                    
                    if(strcmp(VendorName,'DAHENG OPTICS'))
                        w1 = 0.400;
                        w2 = 0.700;
                        Lens_catalog(bz1,7) = num2cell(w1);
                        Lens_catalog(bz1,8) = num2cell(w2);
                    end
                    
                    if(strcmp(VendorName,'EALING'))
                        w1 = 0.400;
                        w2 = 0.700;
                        Lens_catalog(bz1,7) = num2cell(w1);
                        Lens_catalog(bz1,8) = num2cell(w2);
                    end
                    
                    if(strcmp(VendorName,'EDMUND OPTICS'))
                        name = char(LensCatalogs.GetResult(lens-1).LensName); 
                        if(strncmp(name,'121',3))
                            w1 = -1;
                            w2 = -1;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strncmp(name,'35',2))
                            w1 = 0.345;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strncmp(name,'473',3)||strncmp(name,'859',3))
                            w1 = 0.800;
                            w2 = 1.550;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strncmp(name,'659',3)||strncmp(name,'858',3))
                            w1 = 0.345;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        else
                            w1 = 0.400;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        end
                        
                    end
                    
                    if(strcmp(VendorName,'LINOS PHOTONICS'))
                        name = char(LensCatalogs.GetResult(lens-1).LensName); 
                        if(strncmp(name,'033',3))
                            w1 = -1;
                            w2 = -1;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        else
                            w1 = 0.400;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        end
                    end
                    
                    if(strcmp(VendorName,'NEWPORT CORP'))

                        w1 = 0.400;
                        w2 = 0.700;
                        Lens_catalog(bz1,7) = num2cell(w1);
                        Lens_catalog(bz1,8) = num2cell(w2);
                    end
                    
                    if(strcmp(VendorName,'OPTOSIGMA'))

                        w1 = 0.400;
                        w2 = 0.700;
                        Lens_catalog(bz1,7) = num2cell(w1);
                        Lens_catalog(bz1,8) = num2cell(w2);
                    end
                    
                    if(strcmp(VendorName,'ROSS OPTICAL'))
                        name = char(LensCatalogs.GetResult(lens-1).LensName); 
                        if(strfind(name,'L-DLA'))
                            w1 = 0.828;
                            w2 = 0.832;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        else
                            w1 = 0.400;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        end
                        
                    end
                    
                    if(strcmp(VendorName,'THORLABS'))
                        name = char(LensCatalogs.GetResult(lens-1).LensName);
                        focal = LensCatalogs.GetResult(lens-1).EFL;
                        
                        if(strfind(name,'ACA'))
                            w1 = -1;
                            w2 = -1;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strfind(name,'-A'))
                            w1 = 0.400;
                            w2 = 0.700;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strfind(name,'-B'))
                            w1 = 0.600;
                            w2 = 1.050;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        elseif(strfind(name,'-C'))
                            w1 = 1.050;
                            w2 = 1.620;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        else
                            Lens_catalog(bz1,7) = num2cell(-1);
                            Lens_catalog(bz1,8) = num2cell(-1);
                        end
                        
                        if(focal==180)
                            w1 = -1;
                            w2 = -1;
                            Lens_catalog(bz1,7) = num2cell(w1);
                            Lens_catalog(bz1,8) = num2cell(w2);
                        end
                        
                    end
                    
                    
                    Lens_catalog(bz1,1) = num2cell(bz1);
                    Lens_catalog(bz1,2) = vendorname(i);
                    Lens_catalog(bz1,3) = num2cell(LensCatalogs.GetResult(lens-1).EFL);
                    Lens_catalog(bz1,4) = num2cell(LensCatalogs.GetResult(lens-1).EPD);
                    partname = char(LensCatalogs.GetResult(lens-1).LensName);
                    Lens_catalog(bz1,5) = cellstr(partname);
                    
                    if(bz1>1 && lens >= 2)
                        f1 = LensCatalogs.GetResult(lens-1).EFL;
                        f0 = LensCatalogs.GetResult(lens-2).EFL;
                        a1 = LensCatalogs.GetResult(lens-1).EPD;
                        a0 = LensCatalogs.GetResult(lens-2).EPD;
                       if (f1==f0 && a1==a0) 
                           Lens_catalog(bz1,6) = num2cell(k+1);
                       else
                           Lens_catalog(bz1,6) = num2cell(0);
                           k = 0;
                       end
                    else
                        Lens_catalog(bz1,6) = num2cell(0);
                    end
                    bz1 = bz1+1;
                    
                end
                
                Matchlens = 0;
                
            end
            
            
            LensCatalogs.Close();
            %Save and close
            TheSystem.Save();
        
%         end
%         
%     end
    
   end
   
xlswrite('All_Lenses_mid.xlsx', Lens_catalog, 1);   

[catalogN, catalogT] = xlsread('All_Lenses_mid.xlsx');
 
for i = 1:size(catalogN,1)
    
    lens_catalogN1(i,:) = catalogN(i,:);
    lens_catalogT1(i,1) = catalogT(i,1);
                
end  

%set up primary optical system
TheSystem = TheApplication.PrimarySystem;
sampleDir = TheApplication.SamplesDir;

%make a file
testFile = System.String.Concat(sampleDir, 'All_Lenses_mid.zmx');
TheSystem.New(false);
TheSystem.SaveAs(testFile);
    
LensCatalogs = TheSystem.Tools.OpenLensCatalogs();
LensCatalogs.GetAllVendors(); 

for bz1 = 1:size(lens_catalogN1,1)
    
    VendorName1 = cell2mat(lens_catalogT1(bz1,1));
    f1 = lens_catalogN1(bz1,3);
    a1 = lens_catalogN1(bz1,4);
     


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
    Matchlens = LensCatalogs.MatchingLenses;
    Lens_catalog(bz1,9) = num2cell(Matchlens);
 
end
    
    LensCatalogs.Close();
    %Save and close
    TheSystem.Save();
    
    xlswrite(['All_Lenses_mid.xlsx'], Lens_catalog, 1);     
    
    
for bz1 = 1:size(lens_catalogN1,1) 
    Lens_catalog(bz1,10)=num2cell(0);
end
    
    
for bz2 = 1:size(lens_catalogN1,1) 
    
    if(cell2mat(Lens_catalog(bz2,9))==1)
        Lens_catalog(bz2,10)=num2cell(1);
    end
    
    if(cell2mat(Lens_catalog(bz2,9))>1 && cell2mat(Lens_catalog(bz2,10))<1)
        o1 = 1;
        f1 = lens_catalogN1(bz2,3);
        a1 = lens_catalogN1(bz2,4);
        n1 = char(Lens_catalog(bz2,2));
        Lens_catalog(bz2,10)=num2cell(o1);
        for bz3 = bz2+1:size(lens_catalogN1,1)
            if(lens_catalogN1(bz3,3)==f1 && lens_catalogN1(bz3,4)==a1 && strcmp(n1,char(lens_catalogT1(bz3,1))))
                o1 = o1+1;
                Lens_catalog(bz3,10) = num2cell(o1);
                
            end
        end
              
    end
    
    
end
   
    xlswrite(['All_Lenses_app.xlsx'], Lens_catalog, 1); 
    
    flag = 123;
    
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


