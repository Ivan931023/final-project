const API_BASE_URL = "http://127.0.0.1:5000";

document.getElementById("submit-button").addEventListener("click", async () => {
    const stock = document.getElementById("stock").value.trim();
    const start_date = document.getElementById("start_date").value.trim();
    const end_date = document.getElementById("end_date").value.trim();

    if (stock && start_date && end_date) {
        try {
            const response = await fetch(`${API_BASE_URL}/api/analyze_data`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ stock, start: start_date, end: end_date }),
            });

            const data = await response.json();
            if (data.images && data.images.length > 0) {
                const chart = document.getElementById("chart");
                chart.src = data.images[0]; // Display the first chart
                chart.style.display = "block";
            }
        } catch (error) {
            alert(`請求失敗：${error}`);
        }
    } else {
        alert("請填寫所有欄位！");
    }
});

document.getElementById("gpt-button").addEventListener("click", async () => {
    const stock = document.getElementById("stock").value.trim();
    const start_date = document.getElementById("start_date").value.trim();
    const end_date = document.getElementById("end_date").value.trim();
    const question = document.getElementById("gpt-question").value.trim();

    if (stock && start_date && end_date && question) {
        try {
            const response = await fetch(`${API_BASE_URL}/api/send_to_gpt`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    stock_number: stock,
                    start_date: start_date,
                    end_date: end_date,
                    question: question,
                }),
            });

            const data = await response.json();
            document.getElementById("gpt-reply").textContent = data.gpt_reply || "AI 顧問無法提供回答";
        } catch (error) {
            alert(`請求失敗：${error}`);
        }
    } else {
        alert("請輸入所有必要資訊！");
    }
});

document.getElementById("download-button").addEventListener("click", async () => {
    const stock = document.getElementById("stock").value.trim();
    const start_date = document.getElementById("start_date").value.trim();
    const end_date = document.getElementById("end_date").value.trim();

    if (stock && start_date && end_date) {
        try {
            const response = await fetch(`${API_BASE_URL}/api/analyze_data`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ stock, start: start_date, end: end_date }),
            });

            const data = await response.json();
            const excelPath = data.excel_path;

            if (excelPath) {
                const excelResponse = await fetch(`${API_BASE_URL}/api/download_excel?file_path=${excelPath}`);
                const blob = await excelResponse.blob();

                const link = document.createElement("a");
                link.href = window.URL.createObjectURL(blob);
                link.download = "analysis.xlsx";
                link.click();
            }
        } catch (error) {
            alert(`請求失敗：${error}`);
        }
    } else {
        alert("請輸入所有必要資訊！");
    }
});
