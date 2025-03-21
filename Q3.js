async function loadExcel() {
    const response = await fetch("data_ggsheet.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const svgId = "#chart-Q3"; 
    d3.select(svgId).selectAll("*").remove();


    data.forEach(d => {
        d["Thành tiền"] = +d["Thành tiền"] || 0;
        const date = new Date(d["Thời gian tạo đơn"]);
        if (!isNaN(date)) {
            d["Tháng"] = `T${String(date.getMonth() + 1).padStart(2, "0")}`;
        } else {
            d["Tháng"] = "Không xác định";
        }
    });

    const revenueByMonth = d3.rollup(
        data,
        v => d3.sum(v, d => d["Thành tiền"]),
        d => d["Tháng"]
    );

    const aggregatedData = Array.from(revenueByMonth, ([month, revenue]) => ({
        "Tháng": month,
        "Thành tiền": revenue
    })).sort((a, b) => a["Tháng"].localeCompare(b["Tháng"])); 

    const colorScale = d3.scaleOrdinal(d3.schemeCategory10);


    const margin = { top: 50, right: 50, bottom: 50, left: 100 };
    const width = 1200 - margin.left - margin.right;
    const height = 600 - margin.top - margin.bottom; 

    const svg = d3.select(svgId)
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom + 30)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top + 30})`);

    const x = d3.scaleBand()
        .domain(aggregatedData.map(d => d["Tháng"]))
        .range([0, width])
        .padding(0.2);

    const y = d3.scaleLinear()
        .domain([0, d3.max(aggregatedData, d => d["Thành tiền"])])
        .range([height, 0]);

    svg.selectAll(".bar")
        .data(aggregatedData)
        .enter().append("rect")
        .attr("class", "bar")
        .attr("x", d => x(d["Tháng"]))
        .attr("y", d => y(d["Thành tiền"]))
        .attr("width", x.bandwidth())
        .attr("height", d => height - y(d["Thành tiền"]))
        .attr("fill", d => colorScale(d["Tháng"]))
        .on("mouseover", (event, d) => {
            tooltip.style("visibility", "visible")
                .html(`<strong>Tháng:</strong> ${d["Tháng"]}<br>
                       <strong>Doanh số bán:</strong> ${d3.format(",.0f")(d["Thành tiền"])}`);
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
        .attr("x", d => x(d["Tháng"]) + x.bandwidth() / 2)
        .attr("y", d => y(d["Thành tiền"]) - 5)
        .attr("text-anchor", "middle")
        .style("font-size", "12px")
        .style("font-family", "Calibri, sans-serif")
        .text(d => `${d3.format(",.0f")(d["Thành tiền"] / 1e6)} triệu VND`);

    svg.append("g")
        .attr("transform", `translate(0,${height})`)
        .call(d3.axisBottom(x))
        .selectAll("text")
        .style("font-size", "14px")
        .style("font-family", "Calibri, sans-serif");

    svg.append("g")
        .call(d3.axisLeft(y).tickFormat(d => `${d3.format(",.0f")(d / 1e6)}M`))
        .selectAll("text")
        .style("font-size", "14px")
        .style("font-family", "Calibri, sans-serif");

    svg.append("text")
        .attr("x", width / 2)
        .attr("y", -40)
        .attr("text-anchor", "middle")
        .style("font-size", "18px")
        .style("font-weight", "bold")
        .style("font-family", "Calibri, sans-serif")
        .style("fill", "#246ba0")
        .text("Doanh số bán hàng theo Tháng");

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