import re

def main():
    file_path = "index.html"
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    # Find the pillars-container block
    start_tag = '            <div class="pillars-container">'
    end_tag = '            </div>\n        </div>\n    </section>'
    
    # We want to replace everything inside pillars-container
    new_pillars = """            <div class="pillars-container">
                <!-- Pillar 1: Retail -->
                <div class="pillar-row">
                    <div class="pillar-img-container" style="display: grid; grid-template-columns: 1fr 1fr; gap: 4px; height: 350px;">
                        <img src="modern_retail_store.png" alt="HomePro and Retail Store" class="pillar-img">
                        <img src="robinson_retail.png" alt="Robinson Department Store" class="pillar-img">
                    </div>
                    <div class="pillar-content">
                        <div class="pillar-badge">PILLAR 01</div>
                        <h3 class="pillar-name">
                            <span class="lang-en">Retail & Distribution Reach (HomePro & Robinson)</span>
                            <span class="lang-th">เครือข่ายค้าปลีกและการจัดจำหน่าย (โฮมโปร และ โรบินสัน)</span>
                        </h3>
                        <p class="pillar-text">
                            <span class="lang-en">With **HomeProduct Center PCL (HomePro)** as our core association, we command Thailand's undisputed #1 home improvement retail channel. Co-founded by Mr. Manit in 1995 (following his pioneer retail legacy as founder of **Robinson Department Store**), HomePro is listed on the Stock Exchange of Thailand (HMPRO) and operates over 100 branches in Thailand, Malaysia, and Vietnam. We offer a direct pipeline for Japanese premium brands to reach Southeast Asian consumers.</span>
                            <span class="lang-th">ด้วยการเป็นผู้ร่วมก่อตั้ง **โฮมโปร (HMPRO)** บริษัทค้าปลีกสินค้าเกี่ยวกับบ้านอันดับ 1 ของไทย (ต่อยอดจากตํานานผู้บุกเบิกค้าปลีกผู้ร่วมก่อตั้ง **ห้างสรรพสินค้าโรบินสัน** ของคุณมานิต) มีสาขารวมมากกว่า 100 สาขาทั่วประเทศไทย มาเลเซีย และเวียดนาม โฮมโปรจึงเป็นช่องทางการกระจายสินค้าและพันธมิตรค้าปลีกจากญี่ปุ่นที่แข็งแกร่งที่สุดในอาเซียน</span>
                        </p>
                    </div>
                </div>

                <!-- Pillar 2: Real Estate -->
                <div class="pillar-row reverse">
                    <div class="pillar-img-container">
                        <img src="luxury_villa_lake.png" alt="Swan Lake Khao Yai Real Estate" class="pillar-img">
                    </div>
                    <div class="pillar-content">
                        <div class="pillar-badge">PILLAR 02</div>
                        <h3 class="pillar-name">
                            <span class="lang-en">Eco-Luxury Real Estate</span>
                            <span class="lang-th">อสังหาริมทรัพย์ระดับพรีเมียมและเวลเนส</span>
                        </h3>
                        <p class="pillar-text">
                            <span class="lang-en">Under **Elysian Development Co., Ltd.**, we develop boutique projects that combine high-end architecture with deep ecological restoration. Our landmark project, **Swan Lake Residence Khao Yai**, valued at over 3 Billion Baht, transformed a dry landscape into a lush forest by planting over 40,000 trees, allocating 80% of the area to green space and common assets under our founder's philosophy of 'Value First, Profit Follows.' We hold prime land bank assets in Bangkok, Chiang Mai, and Phuket, open for high-end residential, hotel, or retirement resort joint ventures with Japanese developers.</span>
                            <span class="lang-th">ภายใต้ **บริษัท อีลิเชี่ยน ดิเวลลอปเม้นท์ จำกัด** เราเน้นพัฒนาโครงการระดับหรูที่บูรณาการธรรมชาติ โครงการเรือธงอย่าง **สวอนเลค เรสซิเด้นซ์ เขาใหญ่** มูลค่ากว่า 3,000 ล้านบาท ได้รับการฟื้นฟูโดยการปลูกต้นไม้กว่า 40,000 ต้น และจัดสรรพื้นที่โครงการถึง 80% ให้เป็นป่าและส่วนกลาง ตามปรัชญา 'สร้างคุณค่า นำยอดขาย' ของผู้ก่อตั้ง ปัจจุบันมีที่ดินสะสม (Land Bank) พร้อมพัฒนาต่อยอดร่วมทุนกับพันธมิตรญี่ปุ่น</span>
                        </p>
                    </div>
                </div>

                <!-- Pillar 3: FMCG -->
                <div class="pillar-row">
                    <div class="pillar-img-container">
                        <img src="mineral_water_bottle.png" alt="6ty Degrees Mineral Water" class="pillar-img">
                    </div>
                    <div class="pillar-content">
                        <div class="pillar-badge">PILLAR 03</div>
                        <h3 class="pillar-name">
                            <span class="lang-en">FMCG & Premium Beverage (6ty Degrees)</span>
                            <span class="lang-th">กลุ่มสินค้าอุปโภคบริโภคและเวลเนสเบเวอเรจ</span>
                        </h3>
                        <p class="pillar-text">
                            <span class="lang-en">Through **Rare Beverage Co., Ltd.**, we operate a state-of-the-art bottling facility on 18 Rai of land in Chiang Dao, Chiang Mai, as part of a 300-Rai wellness land ecosystem representing an initial investment of **1.0 Billion Baht**. **6ty Degrees** natural mineral water is sourced from a rare geothermal hot spring (60°C) at a depth of 300m+. Certified by three world-renowned testing institutes (**SGS Australia, Intertek UK, and ALS Switzerland**), the facility utilizes advanced high-speed automated bottling lines imported from **Germany** ensuring zero human contamination (850,000 bottles/day capacity). We seek Japanese strategic partners for distribution, OEM/ODM bottling, or wellness beverage joint development.</span>
                            <span class="lang-th">ผ่าน **บริษัท แร่เบฟเวอเรจ จำกัด** เราดำเนินการผลิตน้ำแร่ธรรมชาติแบรนด์ **ซิกตี้ ดีกรี (6ty Degrees)** โดยสร้างโรงงานอัจฉริยะ "RARE" ขนาด 18 ไร่ บนผืนดินกว่า 300 ไร่ ที่อำเภอเชียงดาว จังหวัดเชียงใหม่ ด้วยเงินลงทุนเริ่มแรกกว่า **1,000 ล้านบาท** โดยน้ำแร่ธรรมชาติมีแหล่งกำเนิดจากบ่อน้ำพุร้อนธรรมชาติลึกกว่า 300 เมตร (อุณหภูมิ 60°C) การันตีคุณภาพระดับโลกผ่านการตรวจรับรองจากสถาบันสากล 3 แห่ง (**SGS ออสเตรเลีย, Intertek อังกฤษ, และ ALS สวิตเซอร์แลนด์**) ใช้เทคโนโลยีบรรจุขวดความเร็วสูงไร้สัมผัสจาก **ประเทศเยอรมนี** กำลังการผลิต 850,000 ขวด/วัน เรามองการร่วมทุนกับพันธมิตรญี่ปุ่นทั้งด้านการจัดจำหน่าย, การรับจ้างผลิต (OEM/ODM) หรือการพัฒนาเครื่องดื่มบำรุงสุขภาพร่วมกัน</span>
                        </p>
                    </div>
                </div>

                <!-- Pillar 4: Hospitality & Lifestyle -->
                <div class="pillar-row reverse">
                    <div class="pillar-img-container">
                        <img src="jungceylon_phuket.png" alt="Jungceylon Shopping Mall" class="pillar-img">
                    </div>
                    <div class="pillar-content">
                        <div class="pillar-badge">PILLAR 04</div>
                        <h3 class="pillar-name">
                            <span class="lang-en">Hospitality & Lifestyle Curation (Jungceylon)</span>
                            <span class="lang-th">การบริการและบริหารพื้นที่ไลฟ์สไตล์ (จังซีลอน)</span>
                        </h3>
                        <p class="pillar-text">
                            <span class="lang-en">Led by Ms. Rena Udomkunnatum, we curate and manage premier commercial and retail environments. Our flagship hospitality project includes the space planning, curation, and asset enhancement of the massive **Jungceylon Shopping Center** (200,000+ sqm) in Patong, Phuket. We specialize in creating high-traffic lifestyle destinations, hotel management alliances, and premium brand curation (previously scaling fashion brands like FQ&L nationwide). We welcome joint ventures with Japanese lifestyle brands and hospitality groups.</span>
                            <span class="lang-th">นำโดย คุณรีน่า อุดมคุณธรรม เราบริหารและรังสรรค์พื้นที่เชิงพาณิชย์และไลฟ์สไตล์ระดับพรีเมียม โครงการหลักในกลุ่มการบริการคือการวางแผนพื้นที่และบริหารสิทธิประโยชน์ในศูนย์การค้า **จังซีลอน ป่าตอง ภูเก็ต** (พื้นที่กว่า 200,000 ตร.ม.) รวมถึงการทำแบรนด์แฟชั่นเสื้อผ้า FQ&L ในอดีต และการพัฒนาโครงการโรงแรมบูทิก พร้อมร่วมทุนกับผู้ประกอบการญี่ปุ่นในการพัฒนาแบรนด์ค้าปลีกและไลฟ์สไตล์รูปแบบใหม่</span>
                        </p>
                    </div>
                </div>
            </div>"""

    # Replace pillars block in content
    pattern = re.compile(r'            <div class=\"pillars-container\">.*?            </div>\s*</div>\s*</section>', re.DOTALL)
    
    # We want to replace it
    replacement = new_pillars + "\n        </div>\n    </section>"
    content, count = pattern.subn(replacement, content)
    
    if count > 0:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"[SUCCESS] Replaced {count} occurrences of pillars block in index.html.")
    else:
        print("[ERROR] Pillars block pattern not found in index.html.")

if __name__ == "__main__":
    main()
