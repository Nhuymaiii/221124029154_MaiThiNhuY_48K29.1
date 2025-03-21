async function loadExcel() {
    try {
        const response = await fetch("data_ggsheet.xlsx");
        if (!response.ok) throw new Error("Không thể tải file Excel");
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

        const svgId = "#chart-Q8";
        const svgElement = d3.select(svgId);
        if (svgElement.empty()) throw new Error("Không tìm thấy phần tử SVG với id #chart-Q8");
        svgElement.selectAll("*").remove();


        data.forEach(d => {
            d["Nhóm hàng"] = `[${d["Mã nhóm hàng"]}] ${d["Tên nhóm hàng"]}`;
            const date = new Date(d["Thời gian tạo đơn"]);
            if (!isNaN(date)) {
                const year = date.getFullYear();
                const month = String(date.getMonth() + 1).padStart(2, "0");
                d["Tháng"] = `${year}-${month}`;
                d["Tháng Hiển thị"] = `T${month}`;
            } else {
                d["Tháng"] = "Không xác định";
                d["Tháng Hiển thị"] = "Không xác định";
            }
        });

        const filteredData = data.filter(d => d["Tháng"] !== "Không xác định");

        const totalOrdersByMonth = d3.rollup(
            filteredData,
            v => new Set(v.map(d => d["Mã đơn hàng"])).size,
            d => d["Tháng"]
        );

        const ordersByMonthAndGroup = d3.rollup(
            filteredData,
            v => new Set(v.map(d => d["Mã đơn hàng"])).size,
            d => d["Tháng"],
            d => d["Nhóm hàng"]
        );

        const months = Array.from(totalOrdersByMonth.keys()).sort();
        const displayMonths = months.map(month => `T${month.split("-")[1]}`);
        const groups = [...new Set(filteredData.map(d => d["Nhóm hàng"]))];

        const lineData = groups.map(group => {
            return {
                group: group,
                values: months.map(month => {
                    const totalOrders = totalOrdersByMonth.get(month);
                    const orders = ordersByMonthAndGroup.get(month)?.get(group) || 0;
                    const probability = (orders / totalOrders) * 100 || 0;
                    return { month: month, displayMonth: `T${month.split("-")[1]}`, probability: probability };
                })
            };
        });

        const colorScale = d3.scaleOrdinal(d3.schemeCategory10);



        const tooltip = d3.select("body").append("div")
            .attr("class", "tooltip")
            .style("position", "absolute")
            .style("background", "#fff")
            .style("padding", "5px 10px")
            .style("border", "1px solid #000")
            .style("border-radius", "5px")
            .style("visibility", "hidden")
            .style("font-size", "12px")
            .style("font-family", "Calibri, sans-serif")
            .style("text-align", "left");


        const margin = { top: 50, right: 400, bottom: 50, left: 50 };
        const width = 1400 - margin.left - margin.right;
        const height = 500 - margin.top - margin.bottom;

        const svg = svgElement
            .attr("width", width + margin.left + margin.right)
            .attr("height", height + margin.top + margin.bottom + 30)
            .append("g")
            .attr("transform", `translate(${margin.left},${margin.top + 30})`);

        const x = d3.scaleBand()
            .domain(displayMonths)
            .range([0, width])
            .padding(0.2);

        const y = d3.scaleLinear()
            .domain([0, 100]) 
            .range([height, 0]);

        const line = d3.line()
            .x(d => x(d.displayMonth) + x.bandwidth() / 2)
            .y(d => y(d.probability));

        lineData.forEach(groupData => {
            svg.append("path")
                .datum(groupData.values)
                .attr("fill", "none")
                .attr("stroke", colorScale(groupData.group))
                .attr("stroke-width", 2)
                .attr("d", line);

            svg.selectAll(`.dot-${groupData.group.replace(/[^a-zA-Z0-9]/g, '')}`)
                .data(groupData.values)
                .enter().append("circle")
                .attr("class", `dot-${groupData.group.replace(/[^a-zA-Z0-9]/g, '')}`)
                .attr("cx", d => x(d.displayMonth) + x.bandwidth() / 2)
                .attr("cy", d => y(d.probability))
                .attr("r", 5)
                .attr("fill", colorScale(groupData.group))
                .on("mouseover", (event, d) => {
                    tooltip.style("visibility", "visible")
                        .html(`<strong>Tháng:</strong> ${d.displayMonth}<br>
                               <strong>Nhóm hàng:</strong> ${groupData.group}<br>
                               <strong>Xác suất:</strong> ${d.probability.toFixed(1)}%`);
                })
                .on("mousemove", event => {
                    tooltip.style("left", (event.pageX + 10) + "px")
                           .style("top", (event.pageY - 10) + "px");
                })
                .on("mouseout", () => tooltip.style("visibility", "hidden"));
        });

        svg.append("g")
            .attr("transform", `translate(0,${height})`)
            .call(d3.axisBottom(x))
            .selectAll("text")
            .style("font-size", "12px")
            .style("font-family", "Calibri, sans-serif");

        svg.append("g")
            .call(d3.axisLeft(y).tickFormat(d => `${d.toFixed(0)}%`))
            .selectAll("text")
            .style("font-size", "12px")
            .style("font-family", "Calibri, sans-serif");

        svg.append("text")
            .attr("x", width / 2)
            .attr("y", -20)
            .attr("text-anchor", "middle")
            .style("font-size", "18px")
            .style("font-weight", "bold")
            .style("font-family", "Calibri, sans-serif")
            .style("fill", "#246ba0")
            .text("Xác suất bán hàng của Nhóm hàng theo Tháng");
    } catch (error) {
        console.error("Lỗi khi tải hoặc vẽ biểu đồ:", error);
    }
}

loadExcel();