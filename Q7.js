async function loadExcel() {
    const response = await fetch("data_ggsheet.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const svgId = "#chart-Q7"; 
    d3.select(svgId).selectAll("*").remove();


    data.forEach(d => {
        d["Nhóm hàng"] = `[${d["Mã nhóm hàng"]}] ${d["Tên nhóm hàng"]}`;
    });

    const orderCountByCategory = d3.rollup(
        data,
        v => new Set(v.map(d => d["Mã đơn hàng"])).size,
        d => d["Nhóm hàng"]
    );

    const totalOrders = new Set(data.map(d => d["Mã đơn hàng"])).size;

    const aggregatedData = Array.from(orderCountByCategory, ([category, count]) => ({
        "Nhóm hàng": category,
        "Xác suất bán hàng (%)": (count / totalOrders) * 100
    })).sort((a, b) => b["Xác suất bán hàng (%)"] - a["Xác suất bán hàng (%)"]);

    const colorScale = d3.scaleOrdinal(d3.schemeCategory10);


    const margin = { top: 50, right: 50, bottom: 50, left: 200 };
    const width = 1400 - margin.left - margin.right;
    const height = aggregatedData.length * 30; 

    const svg = d3.select(svgId)
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom + 30)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top + 30})`);

    const x = d3.scaleLinear()
        .domain([0, d3.max(aggregatedData, d => d["Xác suất bán hàng (%)"])])
        .range([0, width]);

    const y = d3.scaleBand()
        .domain(aggregatedData.map(d => d["Nhóm hàng"]))
        .range([0, height])
        .padding(0.2);

    svg.selectAll(".bar")
        .data(aggregatedData)
        .enter().append("rect")
        .attr("class", "bar")
        .attr("x", 0)
        .attr("y", d => y(d["Nhóm hàng"]))
        .attr("width", d => x(d["Xác suất bán hàng (%)"]))
        .attr("height", y.bandwidth())
        .attr("fill", d => colorScale(d["Nhóm hàng"]))
        .on("mouseover", (event, d) => {
            tooltip.style("visibility", "visible")
                .html(`<strong>Nhóm hàng:</strong> ${d["Nhóm hàng"]}<br>
                       <strong>Xác suất bán hàng:</strong> ${d["Xác suất bán hàng (%)"].toFixed(1)}%`);
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
        .attr("x", d => x(d["Xác suất bán hàng (%)"]) + 5)
        .attr("y", d => y(d["Nhóm hàng"]) + y.bandwidth() / 2)
        .attr("dy", "0.35em")
        .style("font-size", "12px")
        .style("font-family", "Calibri, sans-serif")
        .text(d => `${d["Xác suất bán hàng (%)"].toFixed(1)}%`);

    svg.append("g")
        .attr("transform", `translate(0,${height})`)
        .call(d3.axisBottom(x).tickFormat(d => `${d.toFixed(0)}%`))
        .selectAll("text")
        .style("font-size", "12px")
        .style("font-family", "Calibri, sans-serif");

    svg.append("g")
        .call(d3.axisLeft(y))
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
        .text("Xác suất bán hàng theo Nhóm hàng");

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