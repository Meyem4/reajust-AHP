
%   This file is part of DSSLab. Copyright (C) 2020-2021 Meryem Tahri and Haytham Tahri




function varargout = projet(varargin)
% PROJET MATLAB code for AHP.fig
%      PROJET, by itself, creates a new PROJET or raises the existing
%      singleton*.
%
%      H = PROJET returns the handle to a new PROJET or the handle to
%      the existing singleton*.
%
%      PROJET('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PROJET.M with the given input arguments.
%
%      PROJET('Property','Value',...) creates a new PROJET or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before projet_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to projet_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help projet

% Last Modified by GUIDE v2.5 08-Feb-2023 19:56:34



% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @projet_OpeningFcn, ...
                   'gui_OutputFcn',  @projet_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before projet is made visible.
function projet_OpeningFcn(hObject, ~, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to projet (see VARARGIN)

% Choose default command line output for projet
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes projet wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = projet_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes during object creation, after setting all properties.
function uitable3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to uitable3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
filepath = ['Output\'];
% Open the Output directory
winopen(filepath);

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
global filename pathname;
[filename pathname] = uigetfile('*.xlsx', 'Choose FILE to load:','MultiSelect','off');
guidata(hObject, handles);
set(handles.listbox1, 'string', filename);
file = strcat(pathname,filename);
setappdata(0,'pushbutton1',file);
EXPERT = xlsread(getappdata(0,'pushbutton1'));
set(handles.uitable3,'Data',num2cell(EXPERT));


% --------------------------------------------------------------------
function Help_Callback(hObject, eventdata, handles)
% hObject    handle to Help (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
open Help.pdf


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

%PCM = getappdata(0,'pushbutton1');
PCM = xlsread(getappdata(0,'pushbutton1'));;
PCM

%% Step 2: Normalized Matrix

[V,D]=eig(PCM);
X=V(:,1);
somme=sum(X);
jo=X./somme;

[n, m] = size(PCM);
Vpmax = max(eig(PCM));
Vpmax;
CI = (Vpmax-n)/(n-1);
CR = CI/1.12;
CR

CRb=1;
Y=0;
%% Step 3: Calculate Eij = aij* wj/wi
while CRb > 0.1
  for i = 1:n
    for j = 1:m
         C(i,j) = (PCM(i,j)*jo(j)/jo(i));
         C(j,i)= 1/C(i,j);
    end
  end
C;
%% Step 4: Checking high matrice value

M = max(C, [], 'all');
N=n-1;
for h =1:N
  s(1)=max(C, [], 'all');  
    s(h+1)=max(C(C<s(h)));
    if Y==PCM(C==s(h))
        M=s(h+1);      
    end
end    
M;

%% Step 5: Replace max value BY 0 and two corresponding diagonal entries by 2

[x,y]=find(C==M);
[n, m] = size(C);

PCM(C==M)=0;
T = PCM; %%Make another array to fill up
for i = 1:n
    for j = 1:m
        if T(i,j) == 0
            T(j,i) = 0;
            T(i,i) = 2;
            T(j,j) = 2;
        end
    end
end
T;
%% Step 6: two corresponding diagonal entries by 2

[V1,D1]=eig(T);
X1=V1(:,1);
somme1=sum(X1);
joA=X1./somme1;
Scale=[1, 2, 3, 4, 5, 6, 7, 8, 9, 1/9, 1/8, 1/7, 1/6, 1/5, 1/4, 1/3, 1/2];
for i = 1:n
    for j = 1:m
        if PCM(i,j) == 0
            Ag = joA(i)/joA(j);  % calculating wi/wj
        end
    end
end
Ag;
Z=[];
for k=1:17
    Z(k)= abs(Ag - Scale(k));     
end
mi=min(Z);
l=find(Z==mi);
Y=Scale(l);
Z;

for i = 1:n
    for j = 1:m
        if PCM(i,j) == 0
            PCM(i,j) = Y;
            PCM(j,i) = 1/Y;
        end    
    end
end
PCM;
set(handles.uitable1,'Data',PCM);

%% Step 7: Recompute Normalized Matrix/eigen value

[V2,D2]=eig(PCM);
X2=V2(:,1);
somme2=sum(X2);
joB=X2./somme2;
[n, m] = size(PCM);    % calculating CR
VpmaxB = max(eig(PCM));
VpmaxB;
CIb = (VpmaxB-n)/(n-1);

CRb = CIb/1.12;
jo=joB;

CRb
set(handles.text3,'String',CRb)
PCM;
end
Ag

Vpmax

[V,D] = eig(PCM) %produces a diagonal matrix D of eigenvalues and a full matrix V whose columns are the corresponding eigenvectors
W = V(:,1)% In this case
Wn = W./sum(W)

%%%% Weight ranking %%%%
[R,TIEADJ] = tiedrank(-Wn)
T = [Wn R]
set(handles.uitable2,'Data',T);


%%% Plot bar %%%%

x = 1:n; % arbitrary array
y = Wn*[100];

% Create a vector of '%' signs

   pct = char(ones(size(y,1),1)*'%');

% Append the '%' signs after the percentage values

   new_yticks = [char(y),pct];

% 'Reflect the changes on the plot

   set(gca,'yticklabel',new_yticks)

bar(x,y);

ylim([0 100])

xlabel('Criteria');

ylabel('Percentage');

labels = arrayfun(@(value) num2str(value,'%2.2f'),y,'UniformOutput',false);
z = '%';
txt = strcat(labels,z);%concatener y avec %
text(x,y,txt,'HorizontalAlignment','center','VerticalAlignment','bottom')

%%%% Export PCM result and bar chart in Output %%%%
global filename filepath;
expertname = filename;
[name, ~] = strsplit(expertname, '.');
expertname = name{1};
filepath = 'Output\';
ax = gca;
exportgraphics(ax,[filepath, expertname, '_BarChart.pdf'],'ContentType','vector');

box off
result = [filepath, expertname, '.pdf'];
print(result,'-dpdf')
result = [filepath, expertname, '_PCM.xlsx'];
writematrix(PCM,result,'Sheet',1,'Range','A1:AC13')
