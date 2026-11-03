# **1. Executive Summary** 

Renesas Electronics currently navigates a profound strategic inflection point. Over the past five years, the enterprise has aggressively transitioned from a traditional, microcontroller-centric Japanese semiconductor supplier heavily dependent on domestic automotive tier-one manufacturers into a global, mixed-signal, and embedded software platform provider. Through the strategic and highly capitalized acquisitions of Dialog Semiconductor, Celeno, Transphorm, and the transformative $5.9 billion purchase of printed circuit board (PCB) design software leader Altium<sup>1</sup> , Renesas has assembled a formidable portfolio capable of capturing unprecedented value at the intelligent edge. However, severe capital misallocations—most notably the estimated 250 billion JPY impairment loss stemming from the Restructuring Support Agreement (RSA) regarding the Wolfspeed silicon carbide (SiC) wafer supply contract—have forced an abrupt suspension of internal SiC and insulated-gate bipolar transistor (IGBT) research and development<sup>3</sup> . This extraordinary capital destruction dictates a rapid strategic pivot toward medium-voltage gallium nitride (GaN) power architectures and highly integrated embedded processing ecosystems to sustain revenue growth and protect operating margins. The exhaustive analysis conducted herein dictates that Renesas must emphatically increase its commercial focus on Industrial applications. The automotive semiconductor market, while traditionally a stronghold representing nearly half of corporate revenue, is currently experiencing protracted inventory corrections, intensifying pricing pressures from aggressive Chinese market entrants, and a deceleration in global electric vehicle penetration<sup>5</sup> . In stark contrast, the Industrial segment offers superior structural profitability, significantly longer product lifecycles, and a highly fragmented customer base that rewards platform-level integration and insulates the supplier from monopsony pricing power. 

Within the diversified Industrial segment, the analysis identifies Robotics and Physical AI as the single most lucrative and strategically aligned application market for incremental commercial investment over the next fiscal year. While AI data center infrastructure offers intense near-term demand for GaN power devices, the Robotics and Physical AI submarket uniquely leverages the entirety of Renesas’s newly acquired and integrated capabilities. By pairing the latest Arm Cortex-M85 microcontrollers featuring Helium vector extensions, proprietary DRP-AI vision microprocessors, Dialog power management integrated circuits (PMICs), and the Altium electronic design software ecosystem, Renesas can architect an insurmountable economic moat defined by extreme customer switching costs and deeply entrenched software dependencies<sup>7</sup> . 

# **2. Direct Answer to Management Question** 

Management must determine where to focus commercial resources next year to maximize revenue growth while successfully managing business risk. The evidence dictates two definitive strategic decisions. 

First, Renesas must disproportionately increase its commercial focus on Industrial applications 

compared with its automotive and consumer IoT businesses. Automotive market cyclicality, heightened capital intensity, and the strategic retreat from high-voltage SiC development structurally disadvantage Renesas in next-generation electric vehicle powertrains. The Industrial market provides a superior risk-adjusted return profile, benefiting from secular macroeconomic tailwinds in factory automation, edge intelligence, and supply chain reshoring. Crucially, the industrial sector demands the exact medium-voltage GaN, precision analog, and advanced microcontroller components where Renesas currently exhibits technological leadership. Second, Robotics and Physical AI must receive the highest incremental commercial investment next year. This submarket requires complex, multi-component architectural solutions—specifically the fusion of real-time deterministic motor control, low-latency machine vision processing, and high-density power delivery. Robotics and Physical AI perfectly intersects with Renesas’s "Winning Combinations" strategy, maximizing revenue attachment rates per printed circuit board while establishing a recurring revenue dynamic through the Altium software platform, thereby locking in long-term ecosystem dependence. 

# **3. Company and Business Model** 

Renesas Electronics operates as a leading global designer and manufacturer of microcontrollers, microprocessors, analog, mixed-signal, power, and connectivity semiconductors. Historically formed through the consolidation of the semiconductor divisions of Hitachi, Mitsubishi Electric, and NEC, the enterprise has systematically diversified its geographic footprint and technological capabilities to shed its legacy reputation as a slow-moving, domestic-focused supplier<sup>10</sup> . 

The core commercial business model has evolved aggressively from selling discrete, commoditized silicon components to providing subsystem-level reference designs marketed as "Winning Combinations"<sup>11</sup> . This model bundles core processing units with proprietary PMICs, environmental sensors, and wireless connectivity modules, fundamentally altering the commercial lifecycle. By engineering pre-validated hardware combinations, Renesas dramatically accelerates customer time-to-market while significantly increasing its share of wallet per device<sup>13</sup> . Design-ins and design-wins in the industrial sector typically require between 12 and 24 months of rigorous customer qualification testing. However, once a production conversion is achieved, the resulting revenue streams are highly recurrent, spanning industrial product lifecycles of seven to fifteen years with minimal redesign interference. The recent $5.9 billion acquisition of Altium marks a profound and disruptive shift in this traditional semiconductor business model<sup>2</sup> . By integrating Altium’s PCB design software, Renesas is shifting upstream into the electronic design workflow<sup>12</sup> . This maneuver allows the enterprise to capture high-margin, recurring software subscription revenue—Altium reported $263 million in revenue with a 77% recurring profile prior to acquisition—while simultaneously embedding Renesas silicon component libraries directly into the digital workspaces of hardware engineers globally<sup>1</sup> . This creates a frictionless pathway from digital simulation directly to silicon procurement, effectively capturing the customer before competing silicon vendors can engage in the traditional sales cycle<sup>14</sup> . 

# **4. Industrial Business and Reporting Limitations** 

A critical limitation in assessing Renesas’s true industrial market penetration stems from its consolidated financial reporting architecture. The enterprise currently divides its revenue into two primary reporting segments: "Automotive" and "Industrial/Infrastructure/IoT" (I&I)<sup>5</sup> . Treating the entire I&I segment as pure industrial revenue fundamentally distorts market reality, as this segment artificially aggregates heavy factory automation revenue with highly cyclical consumer electronics, smart home devices, and hyperscale data center infrastructure<sup>15</sup> . 

To establish data integrity for commercial resource allocation, the I&I segment must be meticulously disaggregated based on end-market dynamics and historical product mix disclosures. Analysis of global semiconductor end-market consumption suggests that within broad IoT and Infrastructure segments, pure industrial applications—defined strictly as factory automation, robotics, test and measurement, and smart grids—typically represent between 40% and 45% of the total aggregate. 

Applying this methodology to the reported financials reveals a more accurate baseline. In evaluating the FY2025 I&I reported revenue of 671.8 billion JPY<sup>5</sup> , the estimated pure Industrial revenue is calculated at approximately 282 billion to 302 billion JPY. 

|**Revenue**<br>**Category**<br>**(Billion JPY)**|**FY2024 Actual**|**FY2025 Actual**|**Estimated**<br>**Pure**<br>**Industrial**<br>**Share**|**Estimated**<br>**Pure**<br>**Industrial**<br>**Revenue**<br>**(FY2025)**|
|---|---|---|---|---|
|**Consolidated**<br>**Total**|1,348.4|1,321.2|-|-|
|**Automotive**|710.0|639.7|0%|0.0|
|**Segment**|||||
|**I&I Segment**|636.8|671.8|42.5%|285.5|



This distinction is vital for accurate strategic prioritization, as consumer IoT applications carry drastically different margin profiles, pricing power, and inventory cyclicality compared to heavy industrial deployments. Forecasts and resource allocations within this report strictly isolate the pure Industrial addressable market to prevent the misallocation of highly specialized field application engineering (FAE) resources toward lower-margin consumer electronics or hyper-volatile data center deployments. 

# **5. Product Portfolio** 

The Renesas industrial product portfolio represents a highly synergistic ecosystem designed to capture end-to-end signal chain, embedded processing, and power management value across the industrial landscape. 

The microcontroller and microprocessor segments represent the high-margin anchor of the portfolio. The recently launched RA8 series, featuring the Arm Cortex-M85 core with Helium vector extensions, delivers unprecedented deterministic control and edge artificial intelligence inference capabilities without requiring the physical integration of a dedicated neural processing unit<sup>7</sup> . This allows for highly efficient spatial awareness and predictive maintenance algorithms to run directly on the endpoint. For heavier vision-based workloads, the RZ/V series MPUs integrate the proprietary DRP-AI (Dynamically Reconfigurable Processor) architecture. This architecture enables high-speed, low-power machine vision essential for autonomous robotics, real-time defect detection, and complex environmental mapping<sup>16</sup> . 

Following the acquisition of Dialog Semiconductor, Renesas possesses a world-class portfolio of power management integrated circuits. These PMICs are architected to seamlessly interface with Renesas MCUs and MPUs, minimizing required board space, reducing quiescent current, and ensuring strict voltage tolerances in harsh industrial applications<sup>9</sup> . The ability to bundle these PMICs with core processing units creates immense switching costs for industrial customers who rely on Renesas for holistic thermal and power warranties. 

The power semiconductor portfolio is currently undergoing a radical and forced restructuring. The termination of the Wolfspeed SiC supply agreement, resulting in a devastating 250 billion JPY impairment loss, has forced the executive team to suspend internal SiC and IGBT development indefinitely<sup>3</sup> . Consequently, the enterprise is pivoting aggressively toward Gallium Nitride following the acquisition of Transphorm. Transphorm’s normally-off d-mode SuperGaN technology provides superior switching efficiencies for medium-voltage applications, positioning Renesas to dominate the 800V to 48V and 48V to point-of-load conversion markets critical for AI infrastructure and industrial motor drives<sup>19</sup> . 

To complete the endpoint ecosystem, the integration of Celeno provides robust Wi-Fi 6/6E and Bluetooth Low Energy connectivity, which is entirely essential for untethered industrial IoT networks and mobile robotics<sup>22</sup> . Complementing the hardware is the Reality AI software suite, an embedded platform that allows engineers to train machine learning models for predictive maintenance and anomaly detection directly on Renesas MCUs, drastically reducing customer development friction and embedding Renesas deeply into the algorithmic layer of the customer's intellectual property<sup>12</sup> . 

# **6. Application Market Deep Dives** 

The industrial semiconductor landscape is fragmented into highly distinct submarkets, each characterized by unique growth drivers, qualification timelines, and competitive intensities. Factory automation and industrial control represents the traditional core of the industrial semiconductor market. The global factory automation market is projected to expand from 312 billion USD in 2025 to 435 billion USD by 2030, driven by global supply chain reshoring 

initiatives and the relentless demand for predictive maintenance<sup>23</sup> . Semiconductor content is growing exponentially as legacy programmable logic controllers (PLCs) transition to decentralized, edge-intelligent control nodes. Renesas holds strong positioning here with its RZ MPUs and industrial Ethernet solutions. However, the market is highly contested by Texas Instruments and Infineon, both of which possess deeply entrenched, multi-decade relationships with tier-one automation vendors<sup>25</sup> . Qualification timelines in heavy industry are notoriously long—frequently extending up to 36 months—meaning near-term revenue conversion requires displacing highly sticky incumbent architectures, a highly capital-intensive commercial endeavor. Robotics and Physical AI involves the integration of spatial awareness, machine vision, and autonomous mobility into robotic systems. This market is expected to grow at a compound annual growth rate exceeding 30%, reaching over 58 billion USD by 2035<sup>26</sup> . The semiconductor content per robot is massive, requiring synchronized multi-axis motor control, spatial sensors, high-performance vision MPUs, and high-density power delivery. Renesas’s Cortex-M85 MCUs and DRP-AI vision processors are perfectly tailored for this market, offering superior inference-per-watt metrics compared to generic application processors<sup>8</sup> . Customer demand is rapidly shifting from rigid, pre-programmed robotic arms enclosed in safety cages to collaborative robots (cobots) that require real-time environmental adaptation and human interaction. Design-win timelines are moderate at 12 to 18 months, and profitability is exceptional due to the absolute necessity of highly integrated subsystem architectures. AI infrastructure and data centers represent a market experiencing hyper-growth, expanding from 132 billion USD in 2024 to an estimated 371 billion USD by 2030<sup>27</sup> . The primary constraint in modern AI data centers is power density and thermal management. The transition to 800V high-voltage direct current (HVDC) architectures and 48V rack distributions creates a massive addressable market for Renesas’s newly acquired GaN portfolio<sup>21</sup> . While the revenue conversion potential is extremely rapid due to intense, price-inelastic hyperscaler demand, the business carries severe customer concentration risks. Furthermore, it relies heavily on a single technological vector—power management—and fails to pull through the broader MCU, MPU, and embedded software portfolio, limiting long-term strategic entrenchment. 

Energy, electrification, and smart grids require advanced analog sensing, bidirectional power flow control, and highly secure communications protocols. While the market provides long-term macroeconomic stability, the structural retreat from high-voltage SiC development severely impairs Renesas’s ability to compete in grid-scale inverters, high-power energy storage systems, and heavy transmission applications<sup>4</sup> . Consequently, the addressable market for Renesas is largely restricted to endpoint smart meters and lower-voltage residential solar inverters, areas characterized by intense price competition from emerging Asian suppliers leveraging lower manufacturing cost bases. 

Building automation is driven by global energy efficiency mandates and smart city initiatives, requiring low-power MCUs, environmental sensors, and wireless connectivity. Renesas is highly competitive in this space using its RA microcontrollers and Celeno Wi-Fi modules. However, the semiconductor content per device is relatively low, and the technological barriers to entry are minimal. This leads to rapid commoditization and margin compression over the long term, making it unsuitable for disproportionate commercial investment. 

Test, measurement, and instrumentation is a highly specialized submarket requiring ultra-precision analog-to-digital converters, low-noise operational amplifiers, and stable timing solutions. While gross margins in this sector routinely exceed 65%, the total addressable market is exceptionally narrow, and unit volumes are inherently constrained by the niche nature of the end products. Furthermore, Analog Devices and Texas Instruments maintain near-monopolistic control over the highest-precision analog signal chains, making significant market share capture resource-prohibitive for Renesas<sup>29</sup> . 

# **7. Historical Financial Performance** 

An exhaustive analysis of Renesas’s financial performance from FY2021 through the latest available FY2026 data reveals a period of aggressive, acquisition-led expansion followed by recent macroeconomic normalization and significant portfolio restructuring. The enterprise transitioned to International Financial Reporting Standards (IFRS) to align with global peers and successfully integrated multiple large-scale acquisitions, drastically expanding its total addressable market. 

In FY2023, total consolidated revenue reached a peak of 1,469.4 billion JPY, with the Industrial/Infrastructure/IoT segment contributing 764.7 billion JPY<sup>31</sup> . However, FY2024 witnessed a severe cyclical inventory correction across the broader semiconductor industry, bringing total revenue down to 1,348.4 billion JPY, with the I&I segment contracting sharply to 636.8 billion JPY<sup>5</sup> . By the conclusion of FY2025, revenue stabilized at 1,321.2 billion JPY. Notably, the I&I segment recovered slightly to 671.8 billion JPY, surpassing the Automotive segment's 639.7 billion JPY, highlighting the structural resilience of the industrial and infrastructure end-markets relative to the slowing automotive sector<sup>5</sup> . 

|**Financial**<br>**Metric (Billion**<br>**JPY)**|**FY2023 Actual**|**FY2024 Actual**|**FY2025 Actual**|**FY2026**<br>**(Annualized**<br>**Est.)**|
|---|---|---|---|---|
|**Total Revenue**|1,469.4|1,348.4|1,321.2|1,510.4|
|**Automotive**<br>**Revenue**|702.8|710.0|639.7|730.2|
|**I&I Revenue**|764.7|636.8|671.8|772.0|
|**Non-GAAP**<br>**Gross Margin**|57.0%|56.1%|57.6%|58.1%|
|**Non-GAAP**<br>**Operating**<br>**Margin**|34.1%|29.5%|29.3%|32.7%|



The FY2026 annualized estimates are derived from reported first and second-quarter sequential momentum, adjusting for the planned divestment of the Timing business<sup>32</sup> . A defining financial event occurred in FY2025 when Renesas entered a Restructuring Support Agreement with Wolfspeed regarding its massive silicon carbide commitments. This resulted in an expected 250 billion JPY impairment loss on a prior 2 billion USD capacity reservation deposit<sup>3</sup> . This extraordinary capital destruction drastically impacted statutory net income and catalyzed the strategic shift away from SiC toward GaN technologies. To fortify the balance sheet following the $5.9 billion Altium acquisition, Renesas is divesting its Timing business to SiTime for 3 billion USD (structured as 50% cash and 50% stock). Expected to close by the end of 2026, this transaction will generate a non-recurring gain of approximately 1.5 billion USD (234 billion JPY) and provide critical cash infusion for debt deleveraging, returning the enterprise to a stable Net Debt/EBITDA profile<sup>35</sup> . 

# **8. Revenue Generation and Revenue** 

# **Quality** 

Renesas’s revenue quality has structurally improved over the past half-decade. Historically, the enterprise relied heavily on domestic Japanese automotive tier-one suppliers, entities that wielded immense bargaining power and demanded continuous, margin-crushing price downs. The strategic diversification into the highly fragmented global Industrial market has drastically reduced customer concentration risk; currently, no single external customer accounts for more than 10% of consolidated revenue, shielding the enterprise from targeted procurement leverage<sup>37</sup> . 

The transition from discrete component sales to integrated "Winning Combinations" acts as a powerful multiplier for average selling prices (ASPs) and long-term margin capture. By offering pre-engineered system architectures that resolve complex power sequencing and thermal management issues, Renesas elevates customer switching costs from mere component replacement to fundamental system redesign. 

Furthermore, the acquisition of Altium profoundly transforms the quality of Renesas’s revenue by introducing high-margin, recurring software subscription dynamics into a historically cyclical hardware business. As hardware engineers utilize Altium’s cloud-based PCB design platforms, the frictionless integration of Renesas component libraries accelerates design-win conversions. This creates a highly visible, multi-year revenue pipeline that is largely immune to standard semiconductor inventory cyclicality, elevating the predictability of corporate cash flows<sup>1</sup> . 

# **9. Profitability and Unit Economics** 

The unit economics of the Industrial business rely heavily on a highly disciplined, hybrid fab-lite manufacturing model. While Renesas maintains internal mature-node fabrication facilities in Japan to ensure supply continuity for legacy automotive components, it aggressively outsources advanced digital logic, high-performance MPUs, and specialty power geometries to merchant 

foundries, predominantly TSMC<sup>12</sup> . This strategic outsourcing severely limits capital expenditure requirements, reduces depreciation burdens, and completely insulates operating margins during severe industry capacity gluts. 

In FY2025, the I&I segment demonstrated superior structural profitability compared to the Automotive business, generating a non-GAAP gross margin of 60.6% and an operating margin of 27.2%<sup>39</sup> . The gross margin expansion is a direct result of relentless product mix improvements, specifically the higher attachment rates of proprietary, high-margin Dialog PMICs and advanced microcontrollers executing edge AI workloads. However, the operating margin reflects the heavy SG&A and R&D burdens associated with acquiring, integrating, and maintaining the intellectual property portfolios of Dialog, Transphorm, and Altium. Customer acquisition costs in the Industrial space are undeniably high, requiring extensive field application engineering support to navigate complex customer qualification protocols and debug intricate software stacks. Yet, the lifetime value of an industrial design-win heavily justifies this initial operational outlay. Industrial product lifecycles ensure steady, high-margin replenishment orders for a decade or more, generating immense free cash flow long after the initial engineering and R&D expenses have been fully amortized. 

# **10. Competitive Landscape** 

The industrial analog and embedded processing market operates as an unforgiving oligopoly characterized by extreme technical barriers to entry and massive economies of scale. Texas Instruments remains the undisputed global market leader, operating with gross margins routinely exceeding 61.4% and an unmatched 300mm internal manufacturing scale that dictates global analog pricing floors<sup>40</sup> . TI competes ruthlessly on raw cost, catalog breadth, and direct sales reach, but historically struggles with cultivating deep, developer-friendly software ecosystems. Analog Devices dominates the high-precision signal chain, achieving gross margins exceeding 63%<sup>41</sup> . ADI is heavily entrenched in aerospace, defense, and premium test instrumentation, controlling the most lucrative niches of the analog market. Infineon Technologies is the undisputed global leader in power semiconductors, specifically dominating the automotive SiC and industrial IGBT markets<sup>42</sup> . Infineon’s dominance in high-voltage industrial applications severely limits Renesas’s ability to compete in grid-scale energy and heavy industrial traction following the Wolfspeed impairment. STMicroelectronics and NXP remain strong, direct competitors in automotive and industrial microcontrollers. NXP benefits from excellent customer retention in automotive processing<sup>6</sup> , while STMicroelectronics leverages its ubiquitous STM32 ecosystem, which historically rivaled Renesas’s software tooling in ease of use and community adoption. 

To counter TI’s manufacturing scale and ADI’s analog precision, Renesas must compete purely on system-level integration and design friction reduction. The proprietary advantage lies in fusing the Altium PCB design software environment with Reality AI edge algorithms and the Cortex-M85 endpoint hardware. No competitor currently possesses a native, globally dominant PCB design software platform seamlessly linked directly to its silicon product portfolio, providing Renesas with an unassailable workflow monopoly. 

# **11. Market Size and Growth Outlook** 

The global analog integrated circuit market is valued at approximately 89.5 billion USD in 2025 and is projected to grow at a modest 3.0% compound annual growth rate through 2035, reaching 120 billion USD<sup>30</sup> . The broader power semiconductor market is valued at 46.1 billion USD<sup>25</sup> . However, within these broad, mature markets, specific technological vectors such as GaN power semiconductors are expanding at over 11% annually, fueled by the insatiable power density requirements of AI data centers and compact industrial robotics<sup>25</sup> . The total addressable market for Robotics and Physical AI semiconductors is accelerating rapidly. The paradigm shift from rigid, programmatic automation to dynamic, neural-network-driven autonomous movement requires endpoint edge devices to execute highly complex matrix multiplications while operating within strict thermal and power envelopes. This specific subset of the edge AI and industrial robotics market is projected to expand at a CAGR exceeding 25% over the next decade<sup>26</sup> . Crucially, the obtainable market for Renesas is expanding as the enterprise leverages its Altium acquisition to penetrate smaller, agile robotics startups that inherently rely on cloud-based collaboration tools rather than legacy, on-premise engineering software suites. 

# **12. Current Renesas Strategy** 

Renesas’s current strategy is explicitly defined by the fundamental transition from a commoditized silicon vendor to a comprehensive embedded systems platform provider. Official corporate announcements highlight the relentless prioritization of "Intelligence at the Edge," heavily leveraging the DRP-AI architecture in the RZ/V MPUs and the Cortex-M85 cores in the RA8 MCUs to dominate endpoint inference applications<sup>8</sup> . 

The most disruptive strategic maneuver is the ongoing integration of Altium. By launching "Renesas 365, Powered by Altium," the enterprise is actively attempting to collapse the traditional hardware design cycle, bridging the historical gap between component selection, system simulation, and physical PCB layout into a single, seamless digital environment<sup>14</sup> . Simultaneously, the financial fallout from the Wolfspeed restructuring has dictated a forced, yet highly focused, power semiconductor strategy. Abandoning the highly capital-intensive SiC infrastructure race, Renesas is reallocating R&D and capital expenditures entirely toward Transphorm’s d-mode GaN technology<sup>4</sup> . This tactical retreat positions the enterprise to attack the mid-voltage (650V-800V) power conversion markets, specifically targeting hyperscale AI data centers and industrial motor drives where GaN's switching efficiency offers distinct thermal advantages over legacy silicon architectures<sup>28</sup> . 

# **13. Strategic Opportunity Assessment** 

To deploy limited incremental commercial resources effectively, the identified industrial submarkets must be evaluated strictly against Renesas's differentiated capabilities, not merely total market size. 

AI Data Centers offer immediate, high-volume GaN revenue but fundamentally lack long-term 

software stickiness, exposing the business to hyperscaler pricing leverage. Factory Automation provides immense scale but requires displacing deeply entrenched incumbents like TI and Infineon over brutal 36-month qualification cycles, straining FAE resources. Test and Measurement is highly profitable but is structurally walled off by ADI’s precision analog monopoly. 

Robotics and Physical AI, conversely, represents a perfect strategic nexus for Renesas. The market is nascent enough that incumbent compute architectures are not yet fully standardized. The rigorous technical requirements—low-latency vision processing, deterministic motor control, and extreme thermal efficiency—align precisely with Renesas's RA8 MCUs, DRP-AI accelerators, and Dialog PMIC portfolio. Furthermore, the agile engineering teams developing next-generation robotics are the exact target demographic for Altium’s cloud-based design ecosystem, enabling highly efficient customer acquisition. 

# **14. Weighted Commercial Scorecard** 

The following scorecard quantifies the strategic prioritization of the identified industrial submarkets based on Renesas's specific technological and commercial profile. _Note: For Business and Execution Risk, a higher score indicates better risk manageability (i.e., lower actual risk exposure)._ 

|**Criteria**<br>**(Weight**<br>**)**|**AI Data**<br>**Center**<br>**s**|**Factory**<br>**Autom**<br>**ation**|**Roboti**<br>**cs &**<br>**Physic**<br>**al AI**|**Energy**<br>**&**<br>**Smart**<br>**Grids**|**Industri**<br>**al IoT**|**Bldg.**<br>**Autom**<br>**ation**|**Test &**<br>**Measur**<br>**ement**|
|---|---|---|---|---|---|---|---|
|**Revenu**<br>**e**<br>**growth**<br>**(15%)**|9|6|9|5|7|5|4|
|**Addres**<br>**sable**<br>**market**<br>**(15%)**|8|9|8|6|8|6|4|
|**Compe**<br>**titive**<br>**adv.**<br>**(15%)**|6|5|9|3|7|6|3|
|**Profita**<br>**bility**|7|8|8|5|6|5|9|



|**(15%)**||||||||
|---|---|---|---|---|---|---|---|
|**Speed**<br>**of**<br>**conver**<br>**sion**<br>**(10%)**|8|4|6|4|7|7|5|
|**Design-**<br>**win**<br>**visibilit**<br>**y (10%)**|7|8|8|6|6|6|8|
|**Strateg**<br>**ic**<br>**alignm**<br>**ent**<br>**(10%)**|7|8|10|4|8|6|4|
|**Risk**<br>**manag**<br>**eability**<br>**(5%)**|5|7|6|4|7|8|7|
|**Invest**<br>**ment**<br>**feasibil**<br>**ity (5%)**|8|6|8|3|8|8|4|
|**Weight**<br>**ed**<br>**Total**<br>**(100%)**|**7.35**|**6.60**|**8.40**|**4.60**|**7.15**|**5.95**|**5.35**|



Robotics and Physical AI achieves the highest weighted total (8.40). It scores maximum points on Strategic Alignment (10) because it uniquely utilizes the entire depth of the newly constructed Renesas technology stack: Altium for system design, Reality AI for embedded software, Cortex-M85 for endpoint compute, and Dialog for dense power management. The competitive advantage is exceptionally high (9) due to the DRP-AI's superior thermal envelope 

in untethered robotics<sup>7</sup> . Risk manageability is moderate (6) due to the reliance on emerging, highly volatile startup ecosystems rather than legacy incumbents. 

AI Data Centers present a compelling but flawed alternative (7.35). While offering extremely high growth (9) and fast conversion (8) driven by urgent hyperscaler demand for GaN power supplies<sup>28</sup> , the competitive advantage is distinctly lower (6) because Infineon and specialized power players heavily contest this space, and it fundamentally fails to pull through Renesas's high-margin MCU and software portfolio. 

Factory Automation provides a massive TAM (9) and excellent visibility (8), but suffers from terrible speed of conversion (4) due to multi-year design cycles. The competitive advantage is mediocre (5) against TI and Infineon's entrenched safety ecosystems. Finally, Energy and Smart Grids score abysmally on competitive advantage (3) and strategic alignment (4) primarily due to the forced abandonment of high-voltage SiC capabilities following the Wolfspeed RSA collapse<sup>4</sup> . 

# **15. Recommended Industrial Application** 

Robotics and Physical AI is the definitive, unequivocal priority for incremental commercial investment. 

This specific application market provides the optimal risk-adjusted return profile for the enterprise. Unlike AI data centers, which expose Renesas to severe customer concentration risk among a handful of hyperscalers, the robotics ecosystem is broadly distributed across thousands of agile OEMs, academic spin-offs, and industrial system integrators. By targeting this highly fragmented segment, Renesas can monetize its 5.9 billion USD Altium acquisition directly and efficiently. Robotics startups rely on cloud-collaborative design tools to accelerate hardware iterations; aggressively embedding Renesas’s "Winning Combinations" into Altium workflows creates an invisible, structural funnel that defaults engineers into utilizing Renesas silicon prior to interacting with traditional procurement channels<sup>12</sup> . 

Financially, Robotics and Physical AI demands high-margin, heterogeneous compute environments. A single robotic node requires precision motor control (RA MCUs), spatial vision processing (RZ MPUs), and dynamic power regulation (Dialog PMICs). Capturing a single design-win in this space yields a massive revenue multiplier compared to selling standalone microcontrollers into commoditized building automation networks or basic smart meters. 

# **16. Resource Allocation** 

To maximize revenue conversion in Robotics and Physical AI, incremental commercial resources must be deployed specifically to intercept the engineering workflow early in the conceptual design phase. 

|**Commercial Resource**|**Allocation**|**Strategic Justification**|
|---|---|---|
|**Category**|||
|**Field Application**|35%|Complex robotics|
|**Engineering (FAE)**||architectures require<br>intense, localized technical|



|||support for motor control<br>algorithms, ML model<br>deployment (Reality AI),<br>and thermal management<br>optimization.|
|---|---|---|
|**Ecosystem & Software**<br>**Dev.**|25%|Direct integration of<br>Renesas "Winning<br>Combinations" reference<br>designs into the Altium<br>platform; expanding<br>open-source framework<br>support (e.g., ROS2,<br>PyTorch).|
|**Strategic Account**<br>**Development**|20%|Dedicated sales task forces<br>targeting top-tier robotics<br>OEMs and leading physical<br>AI startups, negotiating<br>long-term supply<br>agreements and joint R&D.|
|**Technical Partnerships**|10%|Collaborating with sensor<br>manufacturers (LiDAR,<br>optical) to build<br>pre-validated hardware<br>modules that reduce<br>customer time-to-market.|
|**Product Marketing**|10%|Demand generation<br>focused specifically on the<br>superior inference-per-watt<br>capabilities of the<br>Cortex-M85 Helium<br>extensions and DRP-AI<br>architectures.|
|**Total**|**100%**||



# **17. Revenue Forecast and Scenarios** 

The following FY2027 revenue forecast isolates the pure Industrial addressable market (derived strictly from the I&I reporting segment) and evaluates the impact of the targeted commercial investment in Robotics and Physical AI. The baseline assumption utilizes an estimated FY2026 pure Industrial revenue of 320 billion JPY, anticipating that ASPs will increase due to higher PMIC/MCU attachment rates and that capital allocations from the SiTime divestment will accelerate necessary FAE hiring. 

|**Scenario**|**Est. FY2027 Pure**<br>**Industrial Revenue**|**Gross Margin**<br>**Proxy**|**Expected**<br>**Probability**|
|---|---|---|---|
|**Base Case**|365 billion JPY|61.5%|60%|
|**Bull Case**|410 billion JPY|63.5%|20%|
|**Bear Case**|295 billion JPY|55.0%|20%|



In the Base Case scenario, a steady macroeconomic industrial recovery is observed. The Altium software integration successfully drives a 15% increase in "Winning Combinations" design-wins among mid-tier robotics OEMs. GaN penetration in medium-voltage robotics motors proceeds smoothly, validating the Transphorm acquisition. Commercial return on investment is extremely high, as software ecosystems create sticky, recurring revenue streams that lower long-term customer acquisition costs. 

The Bull Case scenario assumes the rapid, unconstrained proliferation of physical AI in global warehousing and logistics. Hyperscale investments in humanoid robotics create massive pull-through for DRP-AI vision processors. Altium becomes the de facto global standard for AI hardware design, achieving near-ubiquitous Renesas component selection at the simulation layer. Gross margins expand aggressively to 63.5% as software subscriptions and high-margin AI endpoint compute dominate the product mix. 

Conversely, the Bear Case scenario models a global industrial capex freeze exacerbated by escalating geopolitical tensions. Chinese domestic competitors (such as SG Micro or GigaDevice) successfully qualify alternative MCUs, sparking a brutal price war in mid-tier robotics. Gross margins collapse to 55.0% as Renesas is forced to grant severe price concessions merely to maintain fabrication volume and protect market share. 

# **18. Risk Assessment** 

The execution of this commercial strategy carries specific, quantifiable risks that must be actively managed by the executive team to prevent value destruction. 

|**Risk Category**|**Probability**|**Impact**|**Warning**<br>**Indicators**|**Mitigation**<br>**Strategy**|
|---|---|---|---|---|
|**Geopolitical &**<br>**Export**<br>**Control**|High|Severe|Expanded<br>US/Japan<br>restrictions on<br>advanced<br>semiconductor<br>shipments to<br>Chinese AI and<br>robotics firms.|Diversify the<br>customer base<br>heavily toward<br>North<br>American,<br>European, and<br>Indian<br>industrial<br>integrators;<br>localize supply<br>chains where<br>feasible.|
|**Technology**<br>**Substitution**<br>**(RISC-V)**|Medium|High|Rapid adoption<br>of open-source<br>RISC-V<br>architectures<br>displacing<br>premium Arm<br>Cortex MCUs<br>in edge AI.|Accelerate<br>internal<br>RISC-V<br>product lines;<br>ensure the<br>Altium<br>ecosystem<br>supports<br>architecture-ag<br>nostic designs<br>to retain<br>software<br>revenue.|
|**Altium**<br>**Integration**<br>**Failure**|Low|Severe|Engineering<br>talent attrition;<br>failure to<br>successfully<br>embed<br>Renesas IP<br>libraries<br>seamlessly into<br>the cloud<br>platform.|Implement<br>aggressive,<br>multi-year<br>retention<br>packages;<br>establish a<br>dedicated<br>integration<br>management<br>office tracking|



|||||software<br>engagement<br>metrics weekly.|
|---|---|---|---|---|
|**Customer**<br>**Inventory**<br>**Corrections**|Medium|Medium|Rising "Days of<br>Inventory"<br>(DOI) at major<br>channel<br>distributors<br>indicating<br>stalled<br>sell-through.|Shift from a<br>channel<br>booking price<br>basis to a net<br>selling price<br>basis to<br>maintain<br>real-time<br>visibility into<br>actual channel<br>sell-through<sup>45</sup>.|
|**Pricing**<br>**Pressure**<br>**(China)**|High|Medium|Chinese<br>analog<br>suppliers<br>aggressively<br>undercutting<br>PMIC pricing<br>by 30-40%<sup>29</sup>.|Bundle PMICs<br>with proprietary<br>MCUs in<br>complex<br>reference<br>designs to<br>prevent<br>discrete<br>component<br>substitution<br>and obscure<br>unit pricing.|



# **19. Regional Opportunities** 

The geographic deployment of specialized commercial resources must align with the global dispersion of robotics innovation and capital investment. 

North America represents the highest concentration of venture-backed physical AI and robotics startups. The US market is highly receptive to Altium’s cloud-based workflows and demands leading-edge vision processing algorithms (DRP-AI). Intense commercial focus and FAE deployment are required here to capture emerging unicorns before they standardize on competitor architectures. 

Europe, particularly Germany and the Nordics, remains the traditional stronghold of heavy factory automation. While absolute growth is slower than in North America, the region possesses a deep ecosystem of industrial integrators demanding strict functional safety (ISO 26262/IEC 61508) and reliability. Renesas must leverage its automotive-grade quality legacy to 

### penetrate industrial robotics here. 

China is the largest consumer of industrial robotics globally, representing over 30% of total regional revenue historically<sup>15</sup> . However, extreme geopolitical risks and aggressive domestic substitution policies mandate a highly cautious investment approach. Commercial efforts in China must focus solely on high-end, differentiated products that local vendors cannot currently replicate, minimizing exposure to commoditized microcontroller price wars. India is a rapidly emerging manufacturing hub benefiting immensely from global supply chain diversification and "China Plus One" strategies. Renesas’s recent memorandum of understanding with the Indian government to provide Altium software and development boards to academic institutions establishes a vital, long-term talent and ecosystem pipeline that will yield massive commercial dividends over the next decade<sup>46</sup> . 

# **20. Implementation Roadmap** 

To capture dominant market share in Robotics and Physical AI, the commercial organization must execute a highly disciplined 12-month roadmap. 

During the first 90 days, the primary objective is rigorous validation and targeting. The commercial team must extract user analytics from the Altium platform to identify the top 500 emerging robotics and physical AI hardware developers globally. Following this, strategic account teams will be deployed to audit these targets' existing compute and power architectures. Simultaneously, a cross-functional "Robotics Tiger Team" comprising analog, power (GaN), and embedded processing FAEs will be established to unify the commercial approach and eliminate internal product silos. 

Months four through six will focus intensely on engagement and design-in acceleration. The commercial organization will launch targeted technical evaluations, supplying ecosystem partners and selected startups with RA8D1 (Cortex-M85) and RZ/V2H development kits. To lower the barrier to entry, Renesas will provide subsidized access to the Reality AI software suite for rapid machine learning model prototyping directly on Renesas silicon. Crucially, engineering will publish pre-validated "Winning Combination" reference designs specifically engineered for 48V collaborative robot joints, integrating Transphorm GaN, Dialog PMICs, and RA MCUs into a single layout. 

The final phase, spanning months seven through twelve, centers on conversion and production ramp. FAE teams will transition successful technical evaluations into formal, binding design-wins. Supply chain management must allocate priority foundry capacity to ensure unhindered prototyping and initial production ramps for these newly converted accounts, preventing delays that could jeopardize customer launches. Finally, leadership will conduct a comprehensive review of commercial returns, calculating the ASP multiplier achieved by successfully bundling power and analog components with the anchor microcontrollers. 

# **21. Management KPIs** 

Executive management must track specific leading and lagging indicators to evaluate commercial efficacy and course-correct resource deployment rapidly. The primary leading indicator is the Software-Driven Pipeline, measured by the number of unique hardware projects 

initiated within the Altium ecosystem that natively select Renesas MCUs as the primary compute node prior to engaging a sales representative. 

To measure cross-selling success, leadership must track the Attachment Rate, defined as the ratio of proprietary PMICs and GaN power stages sold per MCU/MPU specifically in the robotics segment, targeting greater than 2.5 analog/power ICs per processing node. Operational efficiency will be monitored via the Design-Win Conversion Cycle, tracking the average number of days from an initial customer technical evaluation to a formal production commitment. Finally, to preempt cyclical buildup, Channel Days of Inventory (DOI) must be tracked in real-time specifically for high-end industrial MPUs, while Gross Margin by Submarket is monitored continuously to ensure the Robotics portfolio yields margins exceeding the corporate 58% baseline. 

# **22. CEO and Board-Level Questions** 

Before approving the massive reallocation of commercial FAE resources toward Robotics and Physical AI, the Board of Directors should demand rigorous answers to the following 14 critical questions: 

1. How does the 250 billion JPY impairment from the Wolfspeed SiC agreement fundamentally alter our long-term industrial power roadmap, and does Transphorm's GaN portfolio adequately cover the resulting high-voltage technological gap? 

2. We acquired Altium for an immense $5.9 billion premium; what is the exact, measurable mechanism by which this software platform will drive discrete silicon sales over the next 12 months, rather than merely acting as an independent software vendor? 

3. If Texas Instruments initiates a brutal price war in industrial analog components, how insulated are our "Winning Combinations" from modular substitution by cost-conscious procurement teams? 

4. Does the Arm Cortex-M85 architecture provide a sustainable performance moat against the rapid proliferation of royalty-free RISC-V edge AI cores in the robotics space? 

5. What percentage of the projected robotics revenue is hardware-agnostic software subscription (Altium/Reality AI) versus highly cyclical semiconductor fabrication volume? 

6. By pivoting our elite commercial resources toward emerging physical AI startups, are we risking the degradation of our deeply entrenched relationships with legacy factory automation giants in Japan and Germany? 

7. Can our hybrid fab-lite manufacturing model guarantee the allocation necessary to support hyperscale robotics ramps without requiring massive, margin-dilutive internal CapEx? 

8. With China representing roughly 30% of our revenue historically, how does US-aligned export control legislation impact our ability to ship high-performance DRP-AI processors into the lucrative Chinese robotics market? 

9. Given the rapid, successive integrations of Dialog, Celeno, Transphorm, and Altium, is our internal SG&A structure too bloated to maintain corporate operating margins above the 30% threshold? 

10. Why does our EBITDA multiple structurally lag peers like Analog Devices and NXP, and will this pivot toward industrial software unequivocally correct that market perception? 

11. Is the FAE intensity required to secure highly complex robotics design-wins destroying the unit economics of the initial production ramp? 

12. Following the sale of the Timing business to SiTime, exactly how will the $3 billion in proceeds be deployed to directly support this robotics initiative rather than mere debt deleveraging? 

13. Are we sacrificing immediate, massive, and highly visible revenue in GaN-powered AI data centers by prioritizing the longer-tail, highly fragmented robotics market? 

14. If the Altium ecosystem fails to organically drive silicon attachment as hypothesized, what is the executive contingency plan for recouping the massive acquisition premium without destroying shareholder value? 

# **23. Final Recommendation** 

Renesas Electronics must aggressively and unapologetically increase its commercial focus on Industrial applications to secure durable, high-margin revenue insulated from the structural deceleration and pricing volatility of the automotive sector. Within the expansive industrial landscape, management must direct the highest incremental commercial investment toward Robotics and Physical AI. 

This specific application market fundamentally rewards architectural complexity over component commoditization. It is the singular submarket that fully synthesizes Renesas’s multi-billion-dollar acquisition spree into a cohesive, unassailable value proposition. By utilizing Altium to capture the design engineer’s workflow, Reality AI to simplify machine learning deployment, the Cortex-M85 and DRP-AI for low-latency endpoint compute, and Transphorm GaN paired with Dialog PMICs for high-density power delivery, Renesas transcends the traditional silicon vendor paradigm. 

Executing this strategy will generate a highly visible, recurring revenue pipeline with estimated pure industrial revenues scaling to 365 billion JPY by FY2027 in a base-case scenario. While geopolitical volatility and Chinese analog pricing pressures remain pervasive risks, a highly disciplined, FAE-heavy resource allocation targeting North American and European physical AI ecosystems will yield the greatest sustainable, profitable, and risk-adjusted revenue growth for the enterprise over the next decade. 

# **24. Research Methodology and** 

# **Assumptions** 

The strategic and financial analyses contained within this report are predicated on data and official disclosures available up to the rigorous research cut-off date of November 2, 2026. Financial projections and market growth estimates assume a stabilized macroeconomic environment without severe disruptions to the global semiconductor supply chain or sudden escalations in geopolitical trade embargoes. Currency exchange modeling relies on recent corporate budgeting parameters, anticipating a normalized exchange rate environment to isolate true organic growth from transient foreign exchange tailwinds resulting from the weakened Japanese Yen. Furthermore, data integrity is maintained by mathematically disaggregating the 

reported Industrial/Infrastructure/IoT segment to isolate pure industrial revenues, preventing the strategic distortion caused by hyperscale infrastructure and consumer IoT cyclicality. Calculations surrounding the Wolfspeed impairment and SiTime divestment reflect management's stated expectations and structural timelines as of the fourth quarter of 2026. 

## **Works cited** 

1. Regarding Acquisition of Stock of Altium Limited - Renesas, <u>https://www.renesas.com/en/about/newsroom/regarding-acquisition-stock-altium-li mited</u> 

2. Renesas to Acquire PCB Design Software Leader Altium to Make, <u>https://www.renesas.com/en/about/newsroom/renesas-acquire-pcb-design-softwar e-leader-altium-make-electronics-design-accessible-broader-market</u> 

3. Renesas Announces Expected Loss Resulting from Signing, <u>https://www.renesas.com/en/about/newsroom/renesas-announces-expected-loss-r esulting-signing-restructuring-support-agreement-wolfspeed</u> 

4. Analyst Briefing on Wolfspeed Restructuring Support Agreement, <u>https://www.renesas.com/document/ppt/wolfspeed-restructuring-support-agreeme nt-presentation-and-qa</u> 

5. Renesas Reports Financial Results for the Year Ended December, <u>https://www.renesas.com/document/rep/earnings-report-year-ended-december-31</u> -2025 

6. NXP Semiconductors N.V. (VNX) Business & Moat Analysis (2026), <u>https://koalagains.com/stocks/XETRA/VNX/business-and-moat</u> 

7. RA8P1 - 1GHz Arm Cortex-M85 and Ethos-U55 NPU Based AI, <u>https://www.renesas.com/en/products/ra8p1</u> 

8. Can Your Vision AI Solution Keep Up with Cortex-M85? - Renesas, <u>https://www.renesas.com/en/blogs/can-your-vision-ai-solution-keep-cortex-m85</u> 

9. DA9062 - PMIC Designed for Applications Requiring up to 8.5A, <u>https://www.renesas.com/en/products/da9062</u> 

10. JD_Summer Placements_Renesas Electronics.docx 

11. POWER 2024 Capital Market Day | Renesas, <u>https://www.renesas.com/en/document/ppt/power-2024-capital-market-day</u> 

12. Renesas Strategy and Business Model - Umbrex, <u>https://umbrex.com/resources/company-profiles/renesas-electronics-corp/</u> 

13. Renesas Announces More than 35 Winning Combinations Featuring, <u>https://www.renesas.com/en/about/newsroom/renesas-announces-more-35-winnin g-combinations-featuring-both-dialog-and-renesas-products</u> 

14. Renesas 365 Launches: From Inspired Vision to Reality, <u>https://www.renesas.com/en/blogs/renesas-365-launches-inspired-vision-reality</u> 

15. Renesas Electronics (6723) Revenue Breakdown Overview - 富途牛牛, <u>https://www.futunn.com/en/stock/6723-JP/financials-revenue</u> 

16. Press Release - Renesas - DiscoverEE, 

   - <u>https://www.discoveree.io/morepressrelease/renesas</u> 

17. EEAWARDS2024, <u>https://site.eettaiwan.com/events/eeaward2024/index_en.html</u> 

18. Multi-Channel Power Management ICs (PMICs) - Renesas, <u>https://www.renesas.com/en/products/power-management/multi-channel-power-m anagement-ics-pmics</u> 

19. Renesas Strengthens Power Leadership with New GaN FETs for, <u>https://www.renesas.com/en/about/newsroom/renesas-strengthens-power-leaders hip-new-gan-fets-high-density-power-conversion-ai-data-centers</u> 

20. Renesas Unveils 650 V GaN FETs for Power Conversion - EEPower, <u>https://eepower.com/news/renesas-unveils-650v-gan-fets-for-power-conversion/</u> 

21. What Is 800V HVDC: AI Data Centers Power Architecture - COBTEL, <u>https://www.cobtel.com/info/what-is-nvidia-800v-hvdc-the-power-architectu-10359 4125.html</u> 

22. CL8040 - Wi-Fi 6 Concurrent Dual Band 4T4R PCIe Chip - Renesas, <u>https://www.renesas.com/en/products/cl8040</u> 

23. Industrial Control & Factory Automation Market worth $435.24 billion, <u>https://www.marketsandmarkets.com/core/PressReleases/factory-automation.asp</u> 

24. Factory Automation Market Size & Growth Trends 2031, <u>https://www.mordorintelligence.com/industry-reports/global-factory-automation-an d-industrial-controls-market-industry</u> 

25. Power Semiconductor Market Size, Share & Forecast 2034, <u>https://www.imarcgroup.com/power-semiconductor-market</u> 

26. Physical AI Market Size, Share | Growth [2035], 

   - <u>https://www.astuteanalytica.com/industry-report/physical-ai-market</u> 

27. AI Infrastructure Market Size, Growth and Forecast Report 2030, <u>https://www.techsciresearch.com/report/ai-infrastructure-market/30524.html</u> 

28. Renesas unveils next-gen GaN FETs for data centers - EDN Magazine, <u>https://www.edn.com/renesas-unveils-next-gen-gan-fets-for-data-centers/</u> 

29. Analog Integrated Circuit (IC) Industry Share, Size, Growth & Analysis, <u>https://www.mordorintelligence.com/industry-reports/analog-integrated-circuit-mar ket</u> 

30. Analog Integrated Circuit Market Size, Growth, Forecast Report 2035, <u>https://www.marketresearchfuture.com/reports/analog-integrated-circuit-market-30 156</u> 

31. Financial Report 2023 - Renesas, <u>https://www.renesas.com/en/document/rep/financial-report-2023</u> 

32. 2026 2Q Presentation Material - Renesas, <u>https://www.renesas.com/en/document/ppt/2026-2q-presentation-material</u> 

33. Renesas Q2 2026 slides: revenue surges 25%, operating margin, <u>https://www.investing.com/news/company-news/renesas-q2-2026-slides-revenuesurges-25-operating-margin-hits-33-93CH-4827015</u> 

34. Wolfspeed Nears Bankruptcy; Renesas to Record ¥250B Loss, <u>https://www.smbom.com/news/40327</u> 

35. Renesas Electronics (RNECF) Earnings Date, Time & Estimates, <u>https://www.marketbeat.com/earnings/reports/2026-10-29-renesas-electronics-costock/</u> 

36. SiTime to Acquire Renesas' Timing Business, 

   - <u>https://www.sitime.com/company/newsroom/press-release/sitime-acquire-renesastiming-business</u> 

37. Financial Report 2024 - Renesas, <u>https://www.renesas.com/en/document/rep/financial-report-2024</u> 

38. Financial Report 2022 - Renesas, <u>https://www.renesas.com/en/document/rep/financial-report-2022</u> 

39. Renesas Financial Report 2025, <u>https://www.renesas.com/en/document/rep/financial-report-2025</u> 

40. Texas Instruments (TXN) financial and ratio analysis: Q2 2026 filings, <u>https://sec-api.io/insights/texas-instruments-txn-financial-and-ratio-analysis-q2-202 6-filings-analog-sector-comparison-and-scenarios-to-2029</u> 

41. Analog Devices Reports Strong Fourth Quarter and Fiscal 2025, <u>https://investor.analog.com/news-releases/news-release-details/analog-devices-re ports-strong-fourth-quarter-and-fiscal-2025</u> 

42. Company presentation - Infineon Technologies, <u>https://www.infineon.com/assets/row/public/documents/corporate/fy2026-q2-web-e n-new.pdf</u> 

43. Robotics: Technologies and Global Markets - BCC Research, <u>https://www.bccresearch.com/market-research/engineering/robotics.html</u> 

44. Renesas Delivers New RA8 MCU Group Targeting Graphic Display, <u>https://www.renesas.com/en/about/newsroom/renesas-delivers-new-ra8-mcu-grou p-targeting-graphic-display-solutions-and-voicevision-multimodal-ai</u> 

45. 4Q/FULL-YEAR 2025 PRESENTATION | Renesas, <u>https://www.renesas.com/en/document/ppt/2025-full-year-presentation-material</u> 

46. Renesas Partners with Indian Government to Drive Innovation, <u>https://www.renesas.com/en/about/newsroom/renesas-partners-indian-governmen t-drive-innovation-through-startups-and-industry-academia</u> 

