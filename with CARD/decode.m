%% Decode program for Doublet Selector
% This program is used to decode the lens pair number in Doublet Selector.

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
global best_pair;
global lens1;
global lens2;
lens1 = fix(best_pair/1000000);
mid1 = mod(best_pair,1000000);
order1 = fix(mid1/100000);
mid2 = mod(mid1,10000);
lens2 = fix(mid2/10);
order2 = mod(mid2,10);
