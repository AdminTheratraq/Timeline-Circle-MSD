/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";

import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataViewObjects = powerbi.DataViewObjects;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import * as d3 from 'd3';
import { VisualSettings } from "./settings";
import * as sanitizeHtml from 'sanitize-html';

export interface TimelineData {
    Company: String;
    EventType: string;
    Description: string;
    EventStartDate: Date;
    EventEndDate: Date;
    MoA: String;
    Region: String;
    ProductName: String;
}

export interface Timelines {
    Timeline: TimelineData[];
}

export function logExceptions(): MethodDecorator {
    return (target: Object, propertyKey: string, descriptor: TypedPropertyDescriptor<any>)
        : TypedPropertyDescriptor<any> => {

        return {
            value: function () {
                try {
                    return descriptor.value.apply(this, arguments);
                } catch (e) {
                    // this.svg.append('text').text(e).style("stroke","black")
                    // .attr("dy", "1em");
                    throw e;
                }
            }
        };
    };
}

export function getCategoricalObjectValue<T>(objects: DataViewObjects, index: number, objectName: string, propertyName: string, defaultValue: T): T {
    if (objects) {
        let object = objects[objectName];
        if (object) {
            let property: T = <T>object[propertyName];
            if (property !== undefined) {
                return property;
            }
        }
    }
    return defaultValue;
}

export class Visual implements IVisual {
    private svg: d3.Selection<SVGElement, any, any, any>;
    private margin = { top: 50, right: 40, bottom: 50, left: 40 };
    private settings: VisualSettings;
    private host: IVisualHost;
    private initLoad = false;
    private events: IVisualEventService;
    private xScale: d3.ScaleTime<number, number>;
    private yScale: d3.ScaleLinear<number, number>;
    private gbox: d3.Selection<SVGElement, any, any, any>;
    private colors: any[];

    constructor(options: VisualConstructorOptions) {
        console.log('Visual Constructor', options);
        this.svg = d3.select(options.element).append('svg');
        this.host = options.host;
        this.events = options.host.eventService;
    }

    @logExceptions()
    public update(options: VisualUpdateOptions) {
        console.log('Visual Update ', options);
        this.events.renderingStarted(options);
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        this.svg.selectAll('*').remove();
        let _this = this;
        let vpWidth = (options.viewport.width);
        let vpHeight = (options.viewport.height - 70);
        this.svg.attr('height', vpHeight);
        this.svg.attr('width', vpWidth);

        let gHeight = vpHeight - this.margin.top - this.margin.bottom;
        let gWidth = vpWidth - this.margin.left - this.margin.right;

        let timelineData = Visual.CONVERTER(options.dataViews[0], this.host);
        let minDate, maxDate;

        minDate = new Date(Math.min.apply(null, timelineData.map(d => d.EventStartDate)));
        maxDate = new Date(Math.max.apply(null, timelineData.map(d => d.EventEndDate)));
        minDate = new Date(minDate.getFullYear(), 0, 1);
        maxDate = new Date(maxDate.getFullYear() + 1, 0, 1);

        let months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

        let colors = this.getColors();

        let companyData = timelineData.map(d => d.Company).filter((v, i, self) => self.indexOf(v) === i);
        let companyColorData = companyData.map((d, i) => {
            return {
                company: d,
                color: colors[i]
            };
        });

        this.renderXandYAxis(minDate, maxDate, gWidth, gHeight);

        this.renderTitle(options, gWidth);

        this.defineSVGDefs(companyColorData);

        this.renderXAxisCirclesAndQuarters();

        this.renderTimeRangeLines(timelineData, gHeight);

        this.renderCircles(timelineData, companyColorData);

        this.renderEllipses(companyColorData);

        this.renderText(companyColorData);

        this.handleHyperLinkClick();

        this.renderVisualBorder(options);

        this.events.renderingFinished(options);
    }

    private getColors() {
        return [{
            dark: '#3F5003',
            light: '#D0E987',
            medium: '#AFD045'
        }, {
            dark: '#252D48',
            light: '#81909F',
            medium: '#3B4D64'
        }, {
            dark: '#8D4F0F',
            light: '#D8A26D',
            medium: '#C87825'
        }, {
            dark: '#337779',
            light: '#B2DFE0',
            medium: '#6FCBCC'
        }, {
            dark: '#003366',
            light: '#66ffff',
            medium: '#4791AE'
        }, {
            dark: 'rgba(49, 27, 146,1)',
            light: 'rgba(49, 27, 146,0.2)',
            medium: 'rgba(49, 27, 146,0.5)'
        }, {
            dark: 'rgba(245, 127, 23,1)',
            light: 'rgba(245, 127, 23,0.2)',
            medium: 'rgba(245, 127, 23,0.5)'
        }, {
            dark: 'rgba(183, 28, 28,1)',
            light: 'rgba(183, 28, 28,0.2)',
            medium: 'rgba(183, 28, 28,0.5)'
        }, {
            dark: 'rgba(136, 14, 79,1)',
            light: 'rgba(136, 14, 79,0.2)',
            medium: 'rgba(136, 14, 79,0.5)'
        }, {
            dark: 'rgba(27, 94, 32,1)',
            light: 'rgba(27, 94, 32,0.2)',
            medium: 'rgba(27, 94, 32,0.5)'
        }, {
            dark: 'rgba(255, 0, 0,1)',
            light: 'rgba(255, 0, 0,0.2)',
            medium: 'rgba(255, 0, 0,0.5)'
        }, {
            dark: 'rgba(0, 0, 255,1)',
            light: 'rgba(0, 0, 255,0.2)',
            medium: 'rgba(0, 0, 255,0.5)'
        }, {
            dark: 'rgba(0, 255, 0,1)',
            light: 'rgba(0, 255, 0,0.2)',
            medium: 'rgba(0, 255, 0,0.5)'
        }, {
            dark: 'rgba(94, 89, 27,1)',
            light: 'rgba(94, 89, 27,0.2)',
            medium: 'rgba(94, 89, 27,0.5)'
        }, {
            dark: 'rgba(27, 94, 91,1)',
            light: 'rgba(27, 94, 91,0.2)',
            medium: 'rgba(27, 94, 91,0.5)'
        }, {
            dark: 'rgba(11, 101, 153,1)',
            light: 'rgba(11, 101, 153,0.2)',
            medium: 'rgba(11, 101, 153,0.5)'
        }, {
            dark: 'rgba(11, 45, 153,1)',
            light: 'rgba(11, 45, 153,0.2)',
            medium: 'rgba(11, 45, 153,0.5)'
        }, {
            dark: 'rgba(114, 11, 153,1)',
            light: 'rgba(114, 11, 153,0.2)',
            medium: 'rgba(114, 11, 153,0.5)'
        }, {
            dark: 'rgba(153, 11, 134,1)',
            light: 'rgba(153, 11, 134,0.2)',
            medium: 'rgba(153, 11, 134,0.5)'
        }, {
            dark: 'rgba(249, 5, 134,1)',
            light: 'rgba(249, 5, 134,0.2)',
            medium: 'rgba(249, 5, 134,0.5)'
        }];
    }

    private renderXandYAxis(minDate, maxDate, gWidth, gHeight) {
        let xAxis;
        this.xScale = d3.scaleTime()
            .domain([minDate, maxDate])
            .range([this.margin.left, gWidth]);

        if (this.diff_years(minDate, maxDate) <= 1) {
            xAxis = d3.axisBottom(this.xScale)
                .ticks(d3.timeMonth, 1)
                .tickPadding(20)
                .tickFormat(d3.timeFormat("%b'%y"))
                .tickSize(-10);
        }
        else {
            xAxis = d3.axisBottom(this.xScale)
                .ticks(d3.timeYear, 1)
                .tickPadding(20)
                .tickFormat(d3.timeFormat('%Y'))
                .tickSize(-10);
        }

        let xAxisAllTicks = d3.axisBottom(this.xScale)
            .ticks(d3.timeMonth, 3)
            .tickPadding(20)
            .tickFormat(d3.timeFormat(""))
            .tickSize(10);

        this.yScale = d3.scaleLinear()
            .domain([-100, 100])
            .range([gHeight, this.margin.top]);

        let yAxis = d3.axisLeft(this.yScale);

        let xAxisLineAllTicks = this.svg.append("g")
            .attr("class", "x-axis-line-allticks")
            .attr("transform", "translate(" + (20) + "," + ((gHeight / 2) + 60) + ")")
            .call(xAxisAllTicks);

        let xAxisLine = this.svg.append("g")
            .attr("class", "x-axis-line")
            .attr("transform", "translate(" + (20) + "," + ((gHeight / 2) + 60) + ")")
            .call(xAxis);

        this.svg.append("g")
            .attr("class", "y-axis")
            .call(yAxis).attr('display', 'none');
    }

    private renderTitle(options, gWidth) {
        let gTitle = this.svg.append('g')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', options.viewport.width)
            .attr('height', 50);

        gTitle.append('rect')
            .attr('class', 'chart-header')
            .attr('width', options.viewport.width)
            .attr('height', 35);

        gTitle.append('text')
            .text(this.settings.timeline.title)
            .attr('fill', '#ffffff')
            .attr('font-size', 24)
            .attr('transform', 'translate(' + ((gWidth + 70) / 2 - 104) + ',25)');
    }

    private defineSVGDefs(companyColorData) {
        let svgDefs = this.svg.append('defs');

        companyColorData.forEach((c, i) => {
            let linearGradientTopToBottom = svgDefs.append('linearGradient')
                .attr('x2', '0%')
                .attr('y2', '100%')
                .attr('id', 'linearGradientTopToBottom' + c.company.replace(/ /g, ""));

            linearGradientTopToBottom.append('stop')
                .attr('stop-color', c.color.dark)
                .attr('offset', '0');

            linearGradientTopToBottom.append('stop')
                .attr('stop-color', c.color.light)
                .attr('offset', '1');

            let linearGradientBottomToTop = svgDefs.append('linearGradient')
                .attr('x2', '0%')
                .attr('y2', '100%')
                .attr('id', 'linearGradientBottomToTop' + c.company.replace(/ /g, ""));

            linearGradientBottomToTop.append('stop')
                .attr('stop-color', c.color.light)
                .attr('offset', '0');

            linearGradientBottomToTop.append('stop')
                .attr('stop-color', c.color.dark)
                .attr('offset', '1');
        });
    }

    private renderXAxisCirclesAndQuarters() {
        let year, darkGrey = '#636363', lightGrey = '#868686', color = '#868686';
        this.svg.selectAll('.x-axis-line-allticks .tick').insert('rect')
            .attr('x', 0)
            .attr('y', -25)
            .attr('width', '25%')
            .attr('height', 50)
            .attr('fill', (d: Date, i) => {
                if (i % 4 !== 0) {
                    return color;
                }
                else {
                    if (color === lightGrey) {
                        color = darkGrey;
                    }
                    else {
                        color = lightGrey;
                    }
                    return color;
                }
            });

        this.svg.selectAll('.x-axis-line-allticks .tick line')
            .attr('stroke', '#ffffff')
            .attr('stroke-width', 4);

        this.svg.selectAll('.x-axis-line .tick').insert('circle')
            .attr('cx', 0)
            .attr('cy', 0)
            .attr('r', 27)
            .attr('stroke', '#525252')
            .attr('stroke-width', 4)
            .attr('fill', '#ffffff');

        this.svg.selectAll('.x-axis-line .tick text')
            .attr('y', -5)
            .attr('fill', '#000000').raise();
    }

    private renderTimeRangeLines(timelineData, gHeight) {
        this.svg.selectAll(".line")
            .data(timelineData)
            .enter()
            .append("rect")
            .attr("x", (d: TimelineData, i) => {
                return this.xScale(d.EventStartDate) + 20;
            })
            .attr("width", '8px')
            .attr("y", (d, i) => {
                if (i % 2 === 0) {
                    return this.yScale(-34);
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return this.yScale(54);
                    } else {
                        return this.yScale(14);
                    }
                }
            })
            .attr("height", (d, i) => {
                if (i % 2 === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) {
                        return gHeight - this.yScale(-40);
                    }
                    else {
                        return gHeight - this.yScale(-80);
                    }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return gHeight - this.yScale(-40);
                    }
                    else {
                        return gHeight - this.yScale(-80);
                    }
                }
            })
            .style('fill', (d: TimelineData, i) => {
                if (i % 2 === 0) {
                    return 'url(#linearGradientTopToBottom' + d.Company.replace(/ /g, "") + ')';
                }
                else {
                    return 'url(#linearGradientBottomToTop' + d.Company.replace(/ /g, "") + ')';
                }
            });

        this.svg.selectAll(".line")
            .data(timelineData)
            .enter()
            .append("rect")
            .attr("x", (d: TimelineData, i) => {
                return this.xScale(d.EventEndDate) + 20;
            })
            .attr("width", '8px')
            .attr("y", (d, i) => {
                if (i % 2 === 0) {
                    return this.yScale(-34);
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return this.yScale(54);
                    } else {
                        return this.yScale(14);
                    }
                }
            })
            .attr("height", (d, i) => {
                if (i % 2 === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) {
                        return gHeight - this.yScale(-40);
                    }
                    else {
                        return gHeight - this.yScale(-80);
                    }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return gHeight - this.yScale(-40);
                    }
                    else {
                        return gHeight - this.yScale(-80);
                    }
                }
            })
            .style('fill', (d: TimelineData, i) => {
                if (i % 2 === 0) {
                    return 'url(#linearGradientTopToBottom' + d.Company.replace(/ /g, "") + ')';
                }
                else {
                    return 'url(#linearGradientBottomToTop' + d.Company.replace(/ /g, "") + ')';
                }
            });
    }

    private renderCircles(timelineData, companyColorData) {
        this.gbox = this.svg.selectAll(".box")
            .data(timelineData)
            .enter()
            .append("g")
            .attr('class', (d: TimelineData, i) => {
                if (d.EventType === 'Regulatory') {
                    return 'rect regulatory';
                }
                else if (d.EventType === 'Commercial') {
                    return 'rect commercial';
                }
                else if (d.EventType === 'Clinical Trails') {
                    return 'rect clinical-trails';
                }
            })
            .attr('fill', '#ffffff')
            .attr('transform', (d: TimelineData, i) => {
                let y;
                if ((i % 2) === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) {
                        y = this.yScale(-118);
                    } else {
                        y = this.yScale(-79);
                    }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        y = this.yScale(80);
                    } else {
                        y = this.yScale(40);
                    }
                }
                return 'translate(' + (this.xScale(d.EventStartDate) + 25) + ' ' + y + ')';
            });

        this.gbox.selectAll('g')
            .data((d: any, i) => {
                let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                if (d.EventStartDate.getTime() === d.EventEndDate.getTime() || diff <= 35) {
                    return [d];
                }
                else {
                    return [];
                }
            })
            .enter()
            .append("circle")
            .attr("cx", (d) => {
                let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                if (d.EventStartDate.getTime() !== d.EventEndDate.getTime() && diff <= 35) {
                    return diff / 2;
                }
                else {
                    return 0;
                }
            })
            .attr("cy", 0)
            .attr('r', 40)
            .attr('stroke', (d: TimelineData) => {
                let companyColor = companyColorData.find(c => d.Company === c.company);
                return companyColor ? companyColor.color.light : '#000000';
            })
            .attr('stroke-width', 2)
            .attr('fill', 'rgba(0,0,0,0)');

        this.gbox.selectAll('g')
            .data((d: any, i) => {
                let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                if (d.EventStartDate.getTime() === d.EventEndDate.getTime() || diff <= 35) {
                    return [d];
                }
                else {
                    return [];
                }
            })
            .enter()
            .append('a')
            .append("circle")
            .attr("cx", (d) => {
                let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                if (d.EventStartDate.getTime() !== d.EventEndDate.getTime() && diff <= 35) {
                    return diff / 2;
                }
                else {
                    return 0;
                }
            })
            .attr("cy", 0)
            .attr('r', 45)
            .attr('stroke', (d: TimelineData) => {
                let companyColor = companyColorData.find(c => d.Company === c.company);
                return companyColor ? companyColor.color.medium : '#000000';
            })
            .attr('stroke-width', 4)
            .attr('fill', 'rgba(0,0,0,0)');

    }

    private renderEllipses(companyColorData) {
        this.gbox.selectAll('g')
            .data((d: any, i) => {
                let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                if (d.EventStartDate.getTime() !== d.EventEndDate.getTime() && diff > 35) {
                    return [d];
                }
                else {
                    return [];
                }
            })
            .enter()
            .append('ellipse')
            .attr("cx", (d: TimelineData, i) => {
                let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                return diff / 2;
            })
            .attr("cy", 2)
            .attr("rx", (d: TimelineData, i) => {
                return ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
            })
            .attr("ry", 50)
            .attr('stroke', (d: TimelineData) => {
                let companyColor = companyColorData.find(c => d.Company === c.company);
                return companyColor ? companyColor.color.light : '#000000';
            })
            .attr('stroke-width', 2)
            .attr('fill', 'rgba(0,0,0,0)');

    }

    private renderText(companyColorData) {
        this.gbox.append("foreignObject")
            .html((d) => {
                let companyColor = companyColorData.find(c => d.Company === c.company);
                let color = companyColor ? companyColor.color.medium : '#000000';
                let company = '<div style="color:' + color + ';padding-left: 10px;">' + (d.Company ? sanitizeHtml(d.Company.toString()) : '') + '</div>';
                return company + sanitizeHtml(d.Description);
            })
            .attr('x', (d) => {
                if (d.EventStartDate.getTime() === d.EventEndDate.getTime()) {
                    return -35;
                }
                else {
                    return -20;
                }
            })
            .attr('y', '-50')
            .attr('width', (d) => {
                if (d.EventStartDate.getTime() === d.EventEndDate.getTime()
                    || this.diff_years(d.EventEndDate, d.EventStartDate) < 1) {
                    return 75;
                }
                else {
                    let diff = ((this.xScale(d.EventEndDate) + 25) - (this.xScale(d.EventStartDate) + 25));
                    return diff + diff / 2;
                }
            })
            .attr('height', 70)
            .attr('fill', '#000000')
            .attr('transform', 'translate(0,20)')
            .attr('font-size', 10)
            .attr('font-weight', 'bold');
    }

    private handleHyperLinkClick() {
        let _this = this;
        let baseurl = 'https://strategicanalysisinc.sharepoint.com';
        this.svg.selectAll('foreignObject a')
            .on('click', function (e: Event) {
                e = e || window.event;
                let target: any = e.target || e.srcElement;
                let link = d3.select(this).attr('href');
                if (link.indexOf('http') === -1 || link.indexOf('http') > 0) {
                    link = baseurl + link;
                }
                _this.host.launchUrl(link);
                d3.event.preventDefault();
                return false;
            });
    }

    private renderVisualBorder(options) {
        this.svg.append('rect')
            .attr('class', 'visual-border-rect')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', options.viewport.width)
            .attr('height', options.viewport.height - 70)
            .attr('stroke-width', '2px')
            .attr('stroke', '#333')
            .attr('fill', 'transparent');
    }

    // converter to table data
    public static CONVERTER(dataView: DataView, host: IVisualHost): TimelineData[] {
        let resultData: TimelineData[] = [];
        let tableView = dataView.table;
        let _rows = tableView.rows;
        let _columns = tableView.columns;
        let _companyIndex = -1, _typeIndex = -1, _descIndex = -1, _startDateIndex = -1, _endDateIndex = -1, _moaIndex = -1, _regionIndex, _productIndex;
        for (let ti = 0; ti < _columns.length; ti++) {
            if (_columns[ti].roles.hasOwnProperty("Company")) {
                _companyIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("EventType")) {
                _typeIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Description")) {
                _descIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("EventStartDate")) {
                _startDateIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("EventEndDate")) {
                _endDateIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("MoA")) {
                _moaIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Region")) {
                _regionIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("ProductName")) {
                _productIndex = ti;
            }
        }
        for (let i = 0; i < _rows.length; i++) {
            let row = _rows[i];
            let dp = {
                Company: row[_companyIndex] ? row[_companyIndex].toString() : null,
                EventType: row[_typeIndex] ? row[_typeIndex].toString() : null,
                Description: row[_descIndex] ? row[_descIndex].toString() : null,
                EventStartDate: row[_startDateIndex] ? new Date(Date.parse(row[_startDateIndex].toString())) : null,
                EventEndDate: row[_endDateIndex] ? new Date(Date.parse(row[_endDateIndex].toString())) : null,
                MoA: row[_moaIndex] ? row[_moaIndex].toString() : null,
                Region: row[_regionIndex] ? row[_regionIndex].toString() : null,
                ProductName: row[_productIndex] ? row[_productIndex].toString() : null
            };
            resultData.push(dp);
        }
        return resultData;
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return <VisualSettings>VisualSettings.parse(dataView);
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }

    private diff_years(dt2, dt1) {
        let diff = (dt2.getTime() - dt1.getTime()) / 1000;
        diff /= (60 * 60 * 24);
        return Math.abs(Math.round(diff / 365.25));
    }
}