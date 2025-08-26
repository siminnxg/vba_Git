import requests
import base64
import json
from collections import defaultdict
import webbrowser
import tkinter
import tkinter.ttk
from tkinter import messagebox as msgbox


JIRA_URL = "https://jiradev.nexon.com"  # Jira URL
AUTH = ("simin", "PTvljbMU47PuJGk4oGErDjVMAkgduX2qM67Tbn")  # Jira 사용자 ID 및 API 토큰 
HEADERS = {"Content-Type": "application/json"}  # Json 형식으로 데이터 호출

# 프로젝트 id 추출
def Get_ProjectId(project_name):
    url = f"{JIRA_URL}/rest/api/2/project/{project_name}"
    response = requests.get(url, auth=AUTH, headers=HEADERS)
    if response.status_code == 200:
        data = response.json()
        return data["id"]
        
    else:
        print(f"이슈 조회 실패: {response.status_code}, {response.text}")
        return None


# 사용자가 선택한 프로젝트에서 입력할 field 호출
def get_field(project_key):
    url = f"{JIRA_URL}/rest/api/2/issue/createmeta/{project_key}/issuetypes/1"
    response = requests.get(url, auth=AUTH, headers=HEADERS)
    if response.status_code == 200:
        issue_data = response.json()
        value_list = issue_data["values"]
        
        field_list = defaultdict()
        
        for list in value_list:
            field_list[list["name"]] = list["fieldId"]            
        
        # 특정 필드 제외
        # 설명은 기존 폼이 깨지기 때문에 제외
        del_key = ['설명', '프로젝트', '이슈 유형', '첨부 파일']
        
        for key in del_key:
            # del field_list[key]
            field_list.pop(key)
        
        return field_list
    else:
        print(f"필드 조회 실패: {response.status_code}, {response.text}")
        return None


# 사용자가 입력한 이슈에서 각 필드에 입력되어 있는 값 추출
def get_issue(field_id, issue_key):
    
    # 전역변수 선언
    global Web_URL
        
    url = f"{JIRA_URL}/rest/api/2/issue/{issue_key}"
    response = requests.get(url, auth=AUTH, headers=HEADERS)
    
    if response.status_code == 200:        
        #for list in field_list:
        issue_data = response.json()
        
        for id in field_id:
            value_list = issue_data['fields'][id]              
                        
            if value_list is None:
                continue
            
            # 연결된 이슈 내용 추출
            if id == 'issuelinks':
                if value_list:
                    Web_URL += f"&issuelinks-linktype={value_list[0]['type']['inward']}&issuelinks-issues={value_list[0]['inwardIssue']['key']}&issuelinks-issues"
                continue            
                                                       
            # str 형식 값 추출
            if isinstance(value_list, str):
                Web_URL += f"&{id}={value_list}"
            
            # list 형식 값 추출
            elif isinstance(value_list, list):
                # 공백이 아니면 출력
                if value_list:
                    # 개수가 2개 이상인 경우 반복
                    # 하위에서 id, name 추출
                    if len(value_list) > 1:
                        i = 0
                        while i < len(value_list):
                            if 'id' in value_list[i]: 
                                Web_URL += f"&{id}={value_list[i]['id']}"
                            elif 'name' in value_list[i]:
                                Web_URL += f"&{id}={value_list[i]['name']}"                          
                            else:
                                Web_URL += f"&{id}={value_list[i]}"
                            i += 1
                    else:   
                        if 'id' in value_list[0]: 
                            Web_URL += f"&{id}={value_list[0]['id']}"
                        elif 'name' in value_list[0]:
                            Web_URL += f"&{id}={value_list[0]['name']}"                          
                        else:                     
                            Web_URL += f"&{id}={value_list[0]}"
            
            # id 값 추출
            elif 'id' in value_list:
                Web_URL += f"&{id}={value_list['id']}"
                
                          
    else:
        print(f"이슈 조회 실패: {response.status_code}, {response.text}")
        return None

# 딕셔너리 하위에 딕셔너리가 중첩으로 있는 지 확인
def Search_SubDict(value_list):
    for sub_key in value_list:
        if isinstance(sub_key, dict):
            return sub_key
        else:
            return None


def Btn_TicketCopy():
    
    # 전역변수 선언
    global Web_URL     
    
    Project = combobox.get()
    Ticket = entry_Ticket.get()   
    
    if Project == "":
        msgbox.showinfo("오류","프로젝트명을 입력해주세요.")
        
    elif Ticket == "":
        msgbox.showinfo("오류","이슈 Key를 입력해주세요.")
    
    else:
        # if Project == 'MAGNUM':        
        #     Web_URL = "https://jiradev.nexon.com/secure/CreateIssueDetails!init.jspa?"
        # else:
        #     Web_URL = "http://agjira-stg/secure/CreateIssue.jspa?"
        
        Web_URL = "https://jiradev.nexon.com/secure/CreateIssueDetails!init.jspa?"
        
        # 프로젝트 아이디 추출
        Pid = Get_ProjectId(Project)
        
        # 프로젝트 아이디가 없으면 종료
        if Pid is None:
            msgbox.showinfo("오류","프로젝트명을 확인해주세요.")
            quit()
        else:
            Web_URL += f"pid={Pid}&issuetype=1"
            
        # 각 프로젝트에서 필드 리스트 호출
        field_list = get_field(Pid)        

        # 필드 ID 값만 추출
        field_id = field_list.values()

        if not field_list:
            print ("없음")
        else:
            get_issue(field_id, Ticket)
        
        webbrowser.open(Web_URL)         
   

def Toggle_Frame():
    if frame_Main.winfo_ismapped():  # 현재 표시 중이면
        frame_Main.pack_forget()
    else:  # 숨겨져 있으면 다시 표시
        frame_Main.pack(fill="x")
        

# 이슈 상세 내용 입력하기
def Add_Field():
    
    entries = []
    labels = []
    
    # 프로젝트 선택 여부 확인
    Project = combobox.get()
   
    if Project == "":
        msgbox.showinfo("오류","프로젝트명을 입력해주세요.")
    
    else:
        # 프로젝트 아이디 추출
        Pid = Get_ProjectId(Project)
        # 각 프로젝트에서 필드 리스트 호출
        field_list = get_field(Pid)
        
        if field_list == None:
            quit()
        
        else:
            # 하단 프레임 영역 초기화
            for widget in frame_Issue.winfo_children():
                widget.destroy()
            
            # key 값만 리스트로 변환
            field_name = list(field_list.keys())
            
            # 필드 개수만큼 입력창 추가
            for i in range(len(field_name)):
                labels.append(tkinter.Label(frame_Issue, text=field_name[i], padx=10) )
                labels[i].grid(row=i, column=0)
                entries.append(tkinter.Entry(frame_Issue, width=10))
                entries[i].grid(row=i, column=1)            
        
            
    

###########################################
# 동작 시작

# GUI 설정
window = tkinter.Tk()
window.title("Jira Ticket Info")
window.geometry("500x500+0+0")
window.resizable(False, False)

# 프레임 배치
frame_Main = tkinter.Frame(window, relief="solid", bd="1")
frame_Main.grid(row=0, column=0)

frame_Issue = tkinter.Frame(window, relief="solid", bd="1")
frame_Issue.grid(row=1, column=0)

# 레이블, 버튼 화면 배치
label1 = tkinter.Label(frame_Main, text="프로젝트명")
label1.grid(row=0, column=0)

project_list = ["MAGNUM","DX","MX","DW","RX","MULTIHIT","EXH"]
combobox = tkinter.ttk.Combobox(frame_Main, values=project_list, width=20)
combobox.grid(row=0, column=1)

label2 = tkinter.Label(frame_Main, text="이슈 키")
label2.grid(row=1, column=0)

entry_Ticket = tkinter.Entry(frame_Main, width=20) 
entry_Ticket.grid(row=1, column=1)

button = tkinter.Button(frame_Main, text="이슈 복사하기", command=Btn_TicketCopy)
button.grid(row=1, column=2)

button_Toggle = tkinter.Button(frame_Main, text="이슈 세부 내용 입력하기", command=Add_Field)
button_Toggle.grid(row=2, column=2)

window.mainloop()