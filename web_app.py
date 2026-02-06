import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import datetime
import os

# --- 核心工具：自動處理勾選邏輯 ---
def get_check_context(label, options_map, user_selection):
    result = {}
    for option_text, doc_var in options_map.items():
        if option_text == user_selection:
            result[doc_var] = "☑"
        else:
            result[doc_var] = "☐"
    return result

# 設定網頁標題與佈局
st.set_page_config(page_title="房仲物調表系統", page_icon="🏠")

def main():
    st.title("🏠 房仲物調表 - 快速填寫系統")
    st.markdown("請依序填寫下方資料，完成後點擊最下方的按鈕即可生成 Word 檔。")

    # 檢查範本
    template_name = "template.docx"
    if not os.path.exists(template_name):
        st.error(f"❌ 找不到範本檔案：{template_name}")
        return

    with st.form("survey_form"):
        # ---區塊 1: 基本資料---
        st.subheader("📋 基本資料")
        c1, c2 = st.columns(2)
        with c1:
            listnum = st.text_input("契約委託書編號")
            casename = st.text_input("1. 案名 (必填)", placeholder="例如：住商之星")
            address = st.text_input("2. 物件地址")
            pr = st.text_input("3. 售價 (萬元)")
            # [修正] 修改標題以避免重複 ID
            setprice = st.text_input("設定金額 (萬元)") 
            usearea = st.text_input("使用分區")    
            use = st.text_input("主要用途")    
            
            key_opts = {"公司": "key1", "警衛室": "key2", "洽開發": "key3", "其他:": "key4"}
            sel_key = st.selectbox("建物鑰匙保留", options=list(key_opts.keys()))
                        
        with c2:
            num = st.text_input("編號")
            ve_opts = {
                "1": "ve1", "2": "ve2", "3": "ve3", 
            }
            sel_ve = st.selectbox("開發方式", options=list(ve_opts.keys()))
            community = st.text_input("16. 社區名稱")
            type_opts = {
                "別墅": "t1", "透天": "t2", "電梯華廈": "t3", 
                "套房": "t4", "公寓": "t5", "廠房": "t6", 
                "店面": "t7", "商辦": "t8", "農舍": "t9"
            }
            sel_type = st.selectbox("物件類型", options=list(type_opts.keys()))
            
            state_opts = {"空屋": "state1", "自住": "state2", "出租": "state3"}
            sel_state = st.selectbox("使用現況", options=list(state_opts.keys()))
            
            vd_opts = {"VR": "vd1", "影片": "vd2"}
            sel_vd = st.selectbox("物件是否有影片", options=list(vd_opts.keys()))

            feature = st.text_area("34. 房屋特色", height=150)
            phone = st.text_input("35. 承辦人電話")
          
        # ---區塊 2: 坪數資料---
        st.subheader("📐 坪數資料")
        c1, c2, c3 = st.columns(3)
        with c1:
            totalping = st.text_input("4. 總建坪")
            public_ping = st.text_input("7. 公設坪數")
            addpos = st.text_input("10. 增建位置")
        with c2:
            main_ping = st.text_input("5. 主建物坪數")
            parkingping = st.text_input("8. 汽車位坪數")
            land_ping = st.text_input("31. 土地面積(坪)")
        with c3:
            sub_ping = st.text_input("6. 附屬建物坪數")
            addping = st.text_input("9. 增建坪數")
            land_opts = {"全部持分": "land1", "道路用地": "land2"}
            sel_land = st.selectbox("基地", options=list(land_opts.keys()))
            way = st.text_input("道路坪數")

        # ---區塊 3: 樓層與屋齡---
        st.subheader("🏢 樓層與屋況")
        c1, c2, c3 = st.columns(3)
        with c1:
            totalfloor = st.text_input("11. 總樓層")
            builddate = st.text_input("14. 建築完成日")
            seat = st.text_input("32. 房屋坐向")
        with c2:
            myfloor = st.text_input("12. 位於樓層")
            age = st.text_input("15. 屋齡")
            face = st.text_input("32. 房屋面向")
        with c3:
            underfloor = st.text_input("13. 地下幾層")
            
            car_options = ["坡道平面", "坡道機械", "升降平面", "升降機械", "機械循環", "一樓平面", "無"]
            selected_car_type = st.selectbox("汽車位型式", options=car_options)
            moto = st.text_input("33. 機車車位")

        # ---區塊 4: 格局細節---
        st.subheader("🛋️ 格局配置")
        row1 = st.columns(5)
        room = row1[0].text_input("26. 房")
        hall = row1[1].text_input("27. 廳")
        bath = row1[2].text_input("28. 衛")
        kitchen = row1[3].text_input("29. 廚")
        balcony = row1[4].text_input("30. 陽台")

        row2 = st.columns(2)
        gas_opts = {"桶裝": "gas1", "天然瓦斯": "gas2","電熱器": "gas3","無": "gas4"}
        sel_gas = row2[0].selectbox("瓦斯提供方式", options=list(gas_opts.keys()))
        
        uploaded_file = row2[1].file_uploader("請上傳格局圖 (支援 png, jpg)", type=['png', 'jpg', 'jpeg'])

        # ---區塊 5: 社區與周邊---
        st.subheader("🌳 社區與周邊環境")
        c1, c2 = st.columns(2)
        with c1:
            fee = st.text_input("17. 管理費")
            pay_opts = {"月繳": "pay1", "年繳": "pay2","季繳": "pay3","其他": "pay4"}          
            sel_pay = st.selectbox("管理費繳費方式", options=list(pay_opts.keys()))
            units = st.text_input("19. 同層戶數")
            park = st.text_input("21. 附近公園")
            school = st.text_input("23. 附近學校")
            wi = st.text_input("24. 面寬幾米")
            le = st.text_input("25. 臨路幾米")
        with c2:
            guard_opts = {"有": "guard1", "無": "guard2"}
            sel_guard = st.selectbox("有無警衛", options=list(guard_opts.keys()))
            totalunits = st.text_input("18. 總戶數")
            elevators = st.text_input("20. 電梯數")
            market = st.text_input("22. 附近市場")
            road_opts = {"雙向道": "way1", "單向道": "way2","無尾巷": "way3"}
            sel_road = st.selectbox("巷道狀況", options=list(road_opts.keys()))
            ownduty = st.text_input("增值稅-自用")
            duty = st.text_input("增值稅-一般")

        st.markdown("---")
        submitted = st.form_submit_button("✨ 產生 Word 物調表", type="primary")

    # --- 處理送出後的邏輯 ---
    if submitted:
        if not casename.strip():
            st.error("⚠️ 請輸入「案名」，否則無法產生檔案！")
            return

        context = {
            'listnum': listnum, 'num': num, "casename": casename, "pr": pr,
            "setprice": setprice, "address": address, "community": community,
            "usearea": usearea, "use": use, "phone": phone, "feature": feature,
            "totalping": totalping, "main_ping": main_ping, "sub_ping": sub_ping,
            "public_ping": public_ping, "parkingping": parkingping,
            "land_ping": land_ping, "addping": addping, "addpos": addpos, "way": way,
            "totalfloor": totalfloor, "myfloor": myfloor, "underfloor": underfloor,
            "builddate": builddate, "age": age, "seat": seat, "face": face,
            "room": room, "hall": hall, "bath": bath, "kitchen": kitchen, "balcony": balcony,
            "car_type": selected_car_type, "moto": moto, "fee": fee,
            "totalunits": totalunits, "units": units, "elevators": elevators,
            "park": park, "market": market, "school": school, "wi": wi, "le": le,
            "ownduty": ownduty, "duty": duty,
            "date": datetime.date.today().strftime("%Y/%m/%d")
        }

        # 合併勾選資料
        context.update(get_check_context("Version", ve_opts, sel_ve))
        context.update(get_check_context("Video", vd_opts, sel_vd))
        context.update(get_check_context("Type", type_opts, sel_type))
        context.update(get_check_context("Land", land_opts, sel_land))
        context.update(get_check_context("Guard", guard_opts, sel_guard))
        context.update(get_check_context("State", state_opts, sel_state))
        context.update(get_check_context("Pay", pay_opts, sel_pay))
        context.update(get_check_context("Key", key_opts, sel_key))
        context.update(get_check_context("Road", road_opts, sel_road))
        context.update(get_check_context("Gas", gas_opts, sel_gas))
        
        # ==========================================
        # ⭐ 新增功能：資料預覽區 (Preview)
        # ==========================================
        st.divider() # 分隔線
        st.subheader("🔍 資料核對預覽")
        st.info("請確認下方資料無誤後，再點擊下載按鈕。")

        # 1. 顯示圖片預覽
        if uploaded_file:
            st.image(uploaded_file, caption="格局圖預覽", width=300)
        else:
            st.warning("未上傳格局圖")

        # 2. 顯示重要資料 (用 DataFrame 表格顯示比較整齊)
        import pandas as pd
        
        # 挑選您最想檢查的欄位來顯示
        preview_data = {
            "項目": ["案名", "地址", "總價", "總坪數", "格局", "屋齡", "車位"],
            "內容": [
                casename, 
                address, 
                f"{pr} 萬元", 
                f"{totalping} 坪", 
                f"{room}房 {hall}廳 {bath}衛", 
                f"{age} 年", 
                selected_car_type
            ]
        }
        df = pd.DataFrame(preview_data)
        st.table(df) # 顯示成靜態表格

        # 也可以用 expander 把所有細節藏在裡面，點開才看得到
       
        # ==========================================
        
        try:
            doc = DocxTemplate(template_name)
            
            if uploaded_file:
                image_obj = InlineImage(doc, uploaded_file, width=Mm(50), height=Mm(30))
                context['picture'] = image_obj
            else:
                context['picture'] = "" 

            doc.render(context)

            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            
            output_filename = f"物調表_{casename.strip()}.docx"
            st.success(f"✅ 成功生成！請點擊下方按鈕下載檔案：")
            st.download_button(
                label="📥 點擊下載 Word 檔",
                data=bio,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{e}")
            st.info("請檢查 Word 範本內容是否正確，或確認圖片格式 (jpg/png)。")

if __name__ == "__main__":

    main()

