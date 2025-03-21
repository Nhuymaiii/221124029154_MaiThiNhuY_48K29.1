async function loadExcel() {
    const response = await fetch("data_ggsheet.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const svgId = "#chart-Q4"; 
    d3.select(svgId).selectAll("*").remove();


    const daysOfWeek = ["Chủ nhật", "Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7"];
    data.forEach(d => {
        d["Thành tiền"] = +d["Thành tiền"] || 0;
        const date = new Date(d["Thời gian tạo đơn"]);
        if (!isNaN(date)) {
            d["Ngày trong tuần"] = daysOfWeek[date.getDay()];
            d["Ngày cụ thể"] = date.toDateString();
        } else {
            d["Ngày trong tuần"] = "Không xác định";
            d["Ngày cụ thể"] = "Không xác định";
        }
    });

    const revenueByDay = d3.rollup(
        data.filter(d => d["Ngày trong tuần"] !== "Không xác định"),
        v => ({
            totalRevenue: d3.sum(v, d => d["Thành tiền"]),
            uniqueDays: new Set(v.map(d => d["Ngày cụ thể"])).size
        }),
        d => d["Ngày trong tuần"]
    );

    const aggregatedData = Array.from(revenueByDay, ([day, { totalRevenue, uniqueDays }]) => ({
        "Ngày trong tuần": day,
        "Doanh số trung bình": totalRevenue / uniqueDays
    })).sort((a, b) => daysOfWeek.indexOf(a["Ngày trong tuần"]) - daysOfWeek.indexOf(b["Ngày trong tuần"]));

    const colorScale = d3.scaleOrdinal(d3.schemeCategory10);


    const margin = { top: 50, right: 50, bottom: 50, left: 150 };
    const width = 1200 - margin.left - margin.right;
    const height = 600 - margin.top - margin.bottom;

    const svg = d3.select(svgId)
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom + 30)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top + 30})`);

    const x = d3.scaleBand()
        .domain(aggregatedData.map(d => d["Ngày trong tuần"]))
        .range([0, width])
        .padding(0.2);

    const y = d3.scaleLinear()
        .domain([0, d3.max(aggregatedData, d => d["Doanh số trung bình"])])
        .range([height, 0]);

    svg.selectAll(".bar")
        .data(aggregatedData)
        .enter().append("rect")
        .attr("class", "bar")
        .attr("x", d => x(d["Ngày trong tuần"]))
        .attr("y", d => y(d["Doanh số trung bình"]))
        .attr("width", x.bandwidth())
        .attr("height", d => height - y(d["Doanh số trung bình"]))
        .attr("fill", d => colorScale(d["Ngày trong tuần"]))
        .on("mouseover", (event, d) => {
            tooltip.style("visibility", "visible")
                .html(`<strong>Ngày:</strong> ${d["Ngày trong tuần"]}<br>
                       <strong>Doanh số trung bình:</strong> ${d3.format(",.0f")(d["Doanh số trung bình"])}`);
        })
        .on("mousemove", event => {
            tooltip.style("left", (event.pageX + 10) + "px")
                   .style("top", (event.pageY - 10) + "px");
        })
        .on("mouseout", () => tooltip.style("visibility", "hidden"));

    svg.selectAll(".label")
        .data(aggregatedData)
        .enter().append("text")
        .attr("class", "label")
        .attr("x", d => x(d["Ngày trong tuần"]) + x.bandwidth() / 2)
        .attr("y", d => y(d["Doanh số trung bình"]) - 5)
        .attr("text-anchor", "middle")
        .style("font-size", "12px")
        .style("font-family", "Calibri, sans-serif")
        .text(d => `${d3.format(",.0f")(d["Doanh số trung bình"] / 1e6)} triệu VND`);

    svg.append("g")
        .attr("transform", `translate(0,${height})`)
        .call(d3.axisBottom(x))
        .selectAll("text")
        .style("font-size", "12px")
        .style("font-family", "Calibri, sans-serif");

    svg.append("g")
        .call(d3.axisLeft(y).tickFormat(d => `${d3.format(",.0f")(d / 1e6)}M`))
        .selectAll("text")
        .style("font-size", "12px")
        .style("font-family", "Calibri, sans-serif");

    svg.append("text")
        .attr("x", width / 2)
        .attr("y", -40)
        .attr("text-anchor", "middle")
        .style("font-size", "18px")
        .style("font-weight", "bold")
        .style("font-family", "Calibri, sans-serif")
        .style("fill", "#246ba0")
        .text("Doanh số trung bình theo Ngày trong tuần");

    const tooltip = d3.select("body").append("div")
        .attr("class", "tooltip")
        .style("position", "absolute")
        .style("background", "#fff")
        .style("padding", "5px 10px")
        .style("border", "1px solid #000")
        .style("border-radius", "5px")
        .style("visibility", "hidden")
        .style("font-size", "14px")
        .style("font-family", "Calibri, sans-serif")
        .style("text-align", "left");
}

loadExcel().catch(console.error);