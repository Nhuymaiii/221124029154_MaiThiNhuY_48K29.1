async function loadExcel() {
    const response = await fetch("data_ggsheet.xlsx");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    const svgId = "#chart-Q9"; 
    d3.select(svgId).selectAll("*").remove();

    data.forEach(d => {
        d["Nhóm hàng"] = `[${d["Mã nhóm hàng"]}] ${d["Tên nhóm hàng"]}`;
        d["Mặt hàng"] = `[${d["Mã mặt hàng"]}] ${d["Tên mặt hàng"]}`;
        d["Mã đơn hàng"] = d["Mã đơn hàng"];
    });

    const totalOrdersByGroup = d3.rollup(
        data,
        v => new Set(v.map(d => d["Mã đơn hàng"])).size,
        d => d["Nhóm hàng"]
    );

    const ordersByGroupAndItem = d3.rollup(
        data,
        v => new Set(v.map(d => d["Mã đơn hàng"])).size,
        d => d["Nhóm hàng"],
        d => d["Mặt hàng"]
    );

    const groups = [...new Set(data.map(d => d["Nhóm hàng"]))];
    let transformedData = [];
    groups.forEach(group => {
        const totalOrders = totalOrdersByGroup.get(group) || 0;
        const items = [...new Set(data.filter(d => d["Nhóm hàng"] === group).map(d => d["Mặt hàng"]))];
        items.forEach(item => {
            const orders = ordersByGroupAndItem.get(group)?.get(item) || 0;
            const probability = totalOrders > 0 ? (orders / totalOrders) * 100 : 0;
            transformedData.push({ group, item, probability });
        });
    });

    const topGroups = ["[BOT] Bột", "[SET] Set trà", "[THO] Trà hoa"];
    const categories = [
        ...topGroups,
        ...groups.filter(g => !topGroups.includes(g))
    ];

    const numCols = 3;
    const numRows = 2;
    const margin = { top: 80, right: 50, bottom: 50, left: 350 }; 
    const width = 1700 - margin.left - margin.right; // Tổng chiều rộng SVG giữ nguyên
    const heightPerChart = 200; // Tăng chiều cao mỗi biểu đồ để rõ hơn
    const rowGap = 70; // Tăng khoảng cách giữa các hàng
    const colGap = 250; // Tăng khoảng cách giữa các cột
    const totalHeight = numRows * heightPerChart + (numRows - 1) * rowGap;

    const svg = d3.select(svgId)
        .attr("width", width + margin.left + margin.right)
        .attr("height", totalHeight + margin.top + margin.bottom)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const tooltip = d3.select("body").append("div")
        .attr("class", "tooltip")
        .style("position", "absolute")
        .style("background", "#fff")
        .style("padding", "5px 10px")
        .style("border", "1px solid #000")
        .style("border-radius", "5px")
        .style("visibility", "hidden")
        .style("font-size", "12px") // Tăng kích thước chữ tooltip cho dễ đọc
        .style("font-family", "Calibri, sans-serif")
        .style("text-align", "left");

    const colorScale = d3.scaleOrdinal(d3.schemeCategory10);

    categories.forEach((category, index) => {
        const categoryData = transformedData
            .filter(d => d.group === category)
            .sort((a, b) => b.probability - a.probability);

        const row = Math.floor(index / numCols);
        const col = index % numCols;

        const yOffset = row * (heightPerChart + rowGap);
        const xOffset = col * (width / numCols + colGap);

        const group = svg.append("g")
            .attr("transform", `translate(${xOffset},${yOffset})`);

        const x = d3.scaleLinear()
            .domain([0, d3.max(categoryData, d => d.probability) || 100])
            .range([0, (width - (numCols - 1) * colGap) / numCols]); // Điều chỉnh chiều rộng mỗi biểu đồ

        const y = d3.scaleBand()
            .domain(categoryData.map(d => d.item))
            .range([0, heightPerChart])
            .padding(0.2);

        group.selectAll(".bar")
            .data(categoryData)
            .enter().append("rect")
            .attr("class", "bar")
            .attr("x", 0)
            .attr("y", d => y(d.item))
            .attr("width", d => x(d.probability))
            .attr("height", y.bandwidth())
            .attr("fill", d => colorScale(d.item))
            .on("mouseover", (event, d) => {
                tooltip.style("visibility", "visible")
                    .html(`<strong>Mặt hàng:</strong> ${d.item}<br>
                           <strong>Xác suất:</strong> ${d.probability.toFixed(1)}%`);
            })
            .on("mousemove", event => {
                tooltip.style("left", (event.pageX + 10) + "px")
                       .style("top", (event.pageY - 10) + "px");
            })
            .on("mouseout", () => tooltip.style("visibility", "hidden"));

        group.selectAll(".label")
            .data(categoryData)
            .enter().append("text")
            .attr("class", "label")
            .attr("x", d => x(d.probability) + 5)
            .attr("y", d => y(d.item) + y.bandwidth() / 2)
            .attr("dy", "0.35em")
            .style("font-size", "14px") // Tăng kích thước chữ nhãn cho rõ hơn
            .style("font-family", "Calibri, sans-serif")
            .text(d => `${d.probability.toFixed(1)}%`);

        group.append("g")
            .attr("transform", `translate(0,${heightPerChart})`)
            .call(d3.axisBottom(x).ticks(5).tickFormat(d => `${d.toFixed(0)}%`))
            .selectAll("text")
            .style("font-size", "12px")
            .style("font-family", "Calibri, sans-serif");

        group.append("g")
            .call(d3.axisLeft(y))
            .selectAll("text")
            .style("font-size", "12px")
            .style("font-family", "Calibri, sans-serif")
            .each(function() {
                const text = d3.select(this);
                const label = text.text();
                if (label.length > 20) {
                    text.text(label.substring(0, 17) + "...");
                }
            });

        group.append("text")
            .attr("x", ((width - (numCols - 1) * colGap) / numCols) / 2) // Điều chỉnh vị trí tiêu đề
            .attr("y", -10)
            .attr("text-anchor", "middle")
            .style("font-size", "12px") // Tăng kích thước chữ tiêu đề
            .style("font-weight", "bold")
            .style("font-family", "Calibri, sans-serif")
            .text(category);
    });

    svg.append("text")
        .attr("x", width / 2)
        .attr("y", -40)
        .attr("text-anchor", "middle")
        .style("font-size", "16px") // Tăng nhẹ kích thước chữ tiêu đề chính
        .style("font-weight", "bold")
        .style("font-family", "Calibri, sans-serif")
        .style("fill", "#246ba0")
        .text("Xác suất bán hàng của Mặt hàng theo Nhóm hàng");
}

loadExcel().catch(console.error);