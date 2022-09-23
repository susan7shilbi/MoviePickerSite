const fs = require("fs");
const XLSX = require("xlsx");
const jsontoxml = require("jsontoxml");

const workbook = XLSX.readFile("Book.xlsx");
let worksheets = {};
let SelectedChoices = [];

function Q1(){
    document.getElementsByName("Mood")
        .forEach(radio=>{
            if(radio.checked){
                SelectedChoices[0]= radio.value;
            }
            alert(SelectedChoices[0]);
        })};

// function Q1next(){
//     alert(SelectedChoices[0]);
//     window.open("Q2.html");
// }

function Q2(){
    document.getElementsByName("Ocassion")
        .forEach(radio=>{
            if(radio.checked){
                return SelectedChoices[1]= radio.id;
                }
        })};

function Q3(){
    document.getElementsByName("Genre")
        .forEach(radio=>{
            if(radio.checked){
                return SelectedChoices[2]= radio.id;
            }
        })};

function Q4(){
    document.getElementsByName("rating")
        .forEach(radio=>{
            if(radio.checked){
                return SelectedChoices[3]= radio.id;
            }
        })};

function Q5(){
    document.getElementsByName("old")
        .forEach(radio=>{
            if(radio.checked){
                return SelectedChoices[4]= radio.id;
            }
            worksheets.Sheet1.push({
                "Mood": SelectedChoices[0],
                "Ocassion": SelectedChoices[1],
                "Genre": SelectedChoices[2],
                "Rating": SelectedChoices[3],
                "Old": SelectedChoices[4],
            })
            XLSX.writeFile(workbook, "Book.xlsx")
        })};

document.getElementById("start").onclick=function(){
    location.href = 'Q1.html';
};

// Previous Buttons

document.getElementById("Q1Pre").onclick=function(){
    location.href = 'Main.html';
};

document.getElementById("Q2Pre").onclick=function(){
    location.href = 'Q1.html';
};

document.getElementById("Q3Pre").onclick=function(){
    location.href = 'Q2.html';
};

document.getElementById("Q4Pre").onclick=function(){
    location.href = 'Q3.html';
};

document.getElementById("Q5Pre").onclick=function(){
    location.href = 'Q4.html';
};
