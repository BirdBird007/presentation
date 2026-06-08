// ==========================================================================
// INTERACTIVE SCRIPT - LANG TOGGLE, SLIDESHOW & ALLIANCES
// ==========================================================================

document.addEventListener("DOMContentLoaded", () => {
    
    // --- 1. LANGUAGE TOGGLE LOGIC ---
    const langToggleBtn = document.getElementById("langToggle");
    const body = document.body;

    langToggleBtn.addEventListener("click", () => {
        if (body.classList.contains("lang-en")) {
            body.classList.remove("lang-en");
            body.classList.add("lang-jp");
            langToggleBtn.innerHTML = '<span>EN</span> | <span class="active-lang">JP</span>';
        } else {
            body.classList.remove("lang-jp");
            body.classList.add("lang-en");
            langToggleBtn.innerHTML = '<span class="active-lang">EN</span> | <span>JP</span>';
        }
        // Re-render the collaboration alignment description to match new language
        updateCollabAlignment();
    });

    // --- 2. INTERACTIVE SLIDESHOW LOGIC ---
    const slides = document.querySelectorAll(".web-slide");
    const totalSlides = slides.length;
    let currentSlideIndex = 0; // 0-indexed internally

    const prevBtn = document.getElementById("prevSlideBtn");
    const nextBtn = document.getElementById("nextSlideBtn");
    const progressBar = document.getElementById("progressBar");
    const slideDotsContainer = document.getElementById("slideDots");

    // Generate indicator dots
    for (let i = 0; i < totalSlides; i++) {
        const dot = document.createElement("div");
        dot.classList.add("dot");
        if (i === 0) dot.classList.add("active");
        dot.addEventListener("click", () => goToSlide(i));
        slideDotsContainer.appendChild(dot);
    }
    const dots = document.querySelectorAll(".dot");

    function updateSlideUI() {
        slides.forEach((slide, index) => {
            if (index === currentSlideIndex) {
                slide.classList.add("active");
            } else {
                slide.classList.remove("active");
            }
        });

        // Update Dots
        dots.forEach((dot, index) => {
            if (index === currentSlideIndex) {
                dot.classList.add("active");
            } else {
                dot.classList.remove("active");
            }
        });

        // Update Progress Bar
        const progressPercent = ((currentSlideIndex + 1) / totalSlides) * 100;
        progressBar.style.width = `${progressPercent}%`;
    }

    function goToSlide(index) {
        currentSlideIndex = index;
        updateSlideUI();
    }

    function nextSlide() {
        currentSlideIndex = (currentSlideIndex + 1) % totalSlides;
        updateSlideUI();
    }

    function prevSlide() {
        currentSlideIndex = (currentSlideIndex - 1 + totalSlides) % totalSlides;
        updateSlideUI();
    }

    nextBtn.addEventListener("click", nextSlide);
    prevBtn.addEventListener("click", prevSlide);

    // Keyboard Arrow Keys Navigation
    document.addEventListener("keydown", (e) => {
        // Only trigger if deck viewer is in view
        const rect = document.getElementById("deck-viewer").getBoundingClientRect();
        const inView = rect.top < window.innerHeight && rect.bottom >= 0;
        if (inView) {
            if (e.key === "ArrowRight") {
                nextSlide();
            } else if (e.key === "ArrowLeft") {
                prevSlide();
            }
        }
    });

    // Initialize slide UI
    updateSlideUI();


    // --- 3. DYNAMIC COLLABORATION ALIGNMENT LOGIC ---
    const collabOptions = document.querySelectorAll('input[name="interest"]');
    const alignmentBox = document.getElementById("collabAlignment");

    const alignmentData = {
        real_estate: {
            en: {
                title: "Real Estate & Wellness Alignment",
                description: "Combining Japanese architectural quality and health/senior-care technology with our land bank:",
                bullets: [
                    "Joint venture options on prime land in Khao Yai (adjoining Swan Lake), Chiang Mai, or Phuket.",
                    "Integration of smart eco-technologies, green energy, and elderly-friendly features.",
                    "Direct access to Thailands elite buyers and local high-net-worth real estate networks."
                ]
            },
            jp: {
                title: "不動産・ウェルネス共同事業アライアンス",
                description: "日本の建築技術・デザイン、ヘルスケア・シニアケア技術と、当グループの土地資産を融合します：",
                bullets: [
                    "カオヤイ（スワンレイク隣接地）、チェンマイ、またはプーケットの超一等地における共同開発・JV枠組みの提供",
                    "クリーンエネルギー、省エネ技術、日本のバリアフリー設計（ユニバーサルデザイン）コンセプトの統合",
                    "タイの富裕層顧客および現地プレミアム不動産ネットワークへのダイレクトなアクセス"
                ]
            }
        },
        fmcg: {
            en: {
                title: "6ty Degrees FMCG & OEM Alignment",
                description: "Partnering for high-capacity natural mineral water distribution and product development:",
                bullets: [
                    "Exporting premium geothermal basalt-filtered water to Japans high-end hotel and hospitality sectors.",
                    "OEM/ODM bottling services at our fully automated, 1 Billion+ THB clean-room bottling facility in Chiang Dao.",
                    "Co-developing functional health beverages combining Japanese active ingredients with natural mineral water."
                ]
            },
            jp: {
                title: "6ty Degrees FMCG・OEMアライアンス",
                description: "大容量の天然ミネラルウォーター販売代理店および健康飲料の共同開発における提携：",
                bullets: [
                    "チェンダオの火山岩（バサルト）で濾過されたプレミアム天然温泉水を、日本の高級ホテル・飲食業界（HoReCa）へ輸出",
                    "チェンダオの10億バーツ規模の完全自動化クリーンルーム工場におけるOEM/ODMボトリングサービス",
                    "日本の機能性原材料と当社の天然ミネラルウォーターを組み合わせた、健康ウェルネス飲料の共同開発"
                ]
            }
        },
        retail: {
            en: {
                title: "Retail Brand Import Alignment",
                description: "Leveraging our established modern trade distribution channels to launch Japanese products in Thailand:",
                bullets: [
                    "Direct pipeline into HomePro's 100+ retail branches and online store segments across Thailand.",
                    "Professional marketing localization, logistic warehousing, and retail space management support.",
                    "Strong credibility as co-founders of Robinson and HomePro to easily navigate Thai customs and commercial regulations."
                ]
            },
            jp: {
                title: "小売ブランド輸入アライアンス",
                description: "当グループが確立した強力なモダン・トレード流通網を活用し、日本製品のタイ国内へのローンチを支援します：",
                bullets: [
                    "タイ全土に展開するホームプロ（HomePro）の100以上の実店舗およびオンライン販売網への直通ルート",
                    "現地向けローカライズド・マーケティング、物流倉庫管理、小売ディスプレイ支援の提供",
                    "ロビンソン百貨店およびホームプロの共同創業者としての高い社会的信用による、通関や商取引規制のスムーズな対応"
                ]
            }
        }
    };

    function updateCollabAlignment() {
        const selectedValue = document.querySelector('input[name="interest"]:checked').value;
        const currentLang = body.classList.contains("lang-en") ? "en" : "jp";
        const data = alignmentData[selectedValue][currentLang];

        let bulletsHTML = "";
        data.bullets.forEach(b => {
            bulletsHTML += `<li>${b}</li>`;
        });

        alignmentBox.innerHTML = `
            <div class="alignment-title">${data.title}</div>
            <p>${data.description}</p>
            <ul>${bulletsHTML}</ul>
        `;
    }

    collabOptions.forEach(opt => {
        opt.addEventListener("change", updateCollabAlignment);
    });

    // Initialize alignment box
    updateCollabAlignment();


    // --- 4. FORM SUBMISSION LOGIC ---
    const collabForm = document.getElementById("collabForm");
    const successMsg = document.getElementById("formSuccessMessage");

    collabForm.addEventListener("submit", (e) => {
        e.preventDefault();
        
        // Collect form data (for local demonstration/log)
        const interest = document.querySelector('input[name="interest"]:checked').value;
        const companyName = document.getElementById("companyName").value;
        const email = document.getElementById("contactEmail").value;
        
        console.log("Partnership Inquiry Received:", { interest, companyName, email });

        // Trigger success animation
        collabForm.classList.add("hidden");
        successMsg.classList.remove("hidden");
    });
});
