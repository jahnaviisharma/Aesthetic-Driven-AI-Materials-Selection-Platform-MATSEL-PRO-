# MATSEL PRO
## Aesthetic-Driven AI Materials Selection Platform

### AI-Assisted Material Selection for Engineering, Manufacturing & Product Aesthetics

**Departmental Engineering Project — IIT Ropar**  
**Duration:** January 2026 – May 2026  
**Supervisor:** Dr. Prabir Sarkar  
**Domain:** Materials Engineering · Materials Informatics · AI · Manufacturing · Product Design

---

## Overview

Traditional material selection mainly focuses on engineering properties such as strength, thermal behaviour, and chemical resistance. However, real-world product development also depends on **appearance, surface finish, manufacturing feasibility, cost, sustainability, and user perception**.

MATSEL PRO bridges this gap by connecting:

```text
Product Requirements
        ↓
Aesthetic Intent
        ↓
Material Selection
        ↓
Manufacturing Process
        ↓
Surface Finish
        ↓
Cost Estimation
        ↓
AI-Assisted Design Decision
```

The platform uses a structured database of **850 materials** covering metals, polymers, ceramics, and composites, along with manufacturing processes, cost information, surface characteristics, and aesthetic attributes.

---

##  Key Features

### 🔹 AI Material Recommendation

Uses multi-criteria scoring and dynamic weighting to rank materials based on:

* Engineering requirements
* Aesthetic requirements
* Cost
* Sustainability
* Application
* Production volume

### 🔹 Manufacturing Intelligence

Connects selected materials with compatible manufacturing processes and evaluates achievable surface characteristics such as **Surface Roughness (Ra)** and **Optical Gloss (GU)**.

### 🔹 Cost Estimation

Estimates unit cost using:

**Material + Manufacturing + Surface Treatment + Tooling + Overhead**

and evaluates the effect of production volume on cost.

### 🔹 Vision-Based AI

Analyzes product sketches or renderings using vision-language AI to identify:

* Product characteristics
* Material family
* Surface finish
* Aesthetic intent
* Visual properties

The results are mapped back to the material database to generate recommendations.

### 🔹 Aesthetic Intelligence

Converts design intent into structured aesthetic categories:

`Luxury` · `Functional` · `Organic` · `Industrial` · `Playful` · `Minimal`

The project particularly investigates the relationship between **Ra, GU, surface finish, material perception, and product aesthetics**.

---

##  System Architecture

```text
              USER / DESIGNER
                     ↓
             FRONTEND
       HTML + CSS + JavaScript
                     ↓
              Flask Backend
                     ↓
       ┌─────────────┴─────────────┐
       ↓                           ↓
Material Database            AI / LLM Layer
       ↓                           ↓
       └─────────────┬─────────────┘
                     ↓
       Material + Process + Cost
              Recommendation
```

---

##  Technology Stack

**Materials & Data:** Materials Engineering, Materials Informatics, Pandas, OpenPyXL

**Backend:** Python, Flask, Flask-CORS

**Frontend:** HTML, CSS, JavaScript

**AI:** LLM Integration, Vision AI, Prompt Engineering, OpenRouter API

**Engineering:** Material Selection, Surface Engineering, Manufacturing, Cost Modelling, Multi-Criteria Decision Making

---

##  Project Outcome

MATSEL PRO demonstrates how **engineering data and AI can be combined with product aesthetics** to support material-selection decisions.

Instead of simply asking:

> *"Which material has the required properties?"*

the platform aims to answer:

> **"Which material and manufacturing combination best satisfies the engineering, aesthetic, economic, sustainability, and application requirements?"**

---

##  Project Team

**Jahnavi Sharma** 
B.Tech Mechanical Engineering, IIT Ropar

**Supervisor:** Dr. Prabir Sarkar
Department of Mechanical Engineering, IIT Ropar

---

###  MATSEL PRO

**From Product Aesthetics → Materials → Manufacturing → Cost → AI-Assisted Engineering Decisions**
