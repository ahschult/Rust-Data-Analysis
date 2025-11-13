use calamine::{open_workbook, Reader, Xlsx, Data};
use rust_xlsxwriter::Workbook;
use std::collections::{HashMap, HashSet};
use std::path::Path;

// *** CONFIGURATION ***
const TIME_STANDARDS_FILE: &str = "timestandards.xlsx";
const DATA_FOLDER: &str = "data";
const OUTPUT_FILE: &str = "qualifier_counts.xlsx";

#[derive(Debug, Clone)]
struct MeetResult {
    course: String,
    sex: String,
    age: String,
    event: String,
    time: f64,
    name: String,
}

type StandardKey = (String, String, String); // (sex, age, event)
type AgeGroupStandards = HashMap<String, f64>; // {age_group: qualifying_time}
type EventStandards = HashMap<String, AgeGroupStandards>; // {event: {age: time}}

fn normalize_event_name(event: &str) -> Option<String> {
    if event.trim().is_empty() {
        return None;
    }

    let mut normalized = event.trim().to_string();
    
    // Remove 'm' and ALL spaces
    normalized = normalized.replace('m', "").replace(' ', "");
    
    // Normalize stroke names to abbreviations
    // Full names from data files -> 2-letter abbreviations
    normalized = normalized.replace("Free", "Fr");
    normalized = normalized.replace("Fly", "Bu");  // Butterfly
    normalized = normalized.replace("Back", "Bk");
    normalized = normalized.replace("Breast", "Br");
    normalized = normalized.replace("M.E.", "Me");  // Medley (I.M.)
    normalized = normalized.replace("M.E", "Me");
    normalized = normalized.replace("I.M.", "Me");
    normalized = normalized.replace("I.M", "Me");
    
    // Also handle abbreviated forms that might come from standards file
    normalized = normalized.replace("FL", "Bu");
    
    Some(normalized)
}

fn normalize_age(age: &str) -> String {
    // Remove "&U" suffix if present
    age.trim().replace("&U", "")
}

fn find_best_age_match(athlete_age: &str, available_ages: &[String]) -> Option<String> {
    let athlete_age_num = athlete_age.parse::<i32>().ok()?;
    
    // Convert available ages to numbers
    let mut age_nums: Vec<(i32, String)> = available_ages
        .iter()
        .filter_map(|a| {
            a.parse::<i32>().ok().map(|num| (num, a.clone()))
        })
        .collect();
    
    if age_nums.is_empty() {
        return None;
    }
    
    // Sort by age
    age_nums.sort_by_key(|(num, _)| *num);
    
    // Find exact match first
    if let Some((_, age_str)) = age_nums.iter().find(|(num, _)| *num == athlete_age_num) {
        return Some(age_str.clone());
    }
    
    // If athlete is younger than minimum standard, use the minimum
    if athlete_age_num < age_nums[0].0 {
        return Some(age_nums[0].1.clone());
    }
    
    // If athlete is older than maximum standard, use the maximum
    if athlete_age_num > age_nums.last().unwrap().0 {
        return Some(age_nums.last().unwrap().1.clone());
    }
    
    // Find closest age (shouldn't normally reach here, but just in case)
    let closest = age_nums
        .iter()
        .min_by_key(|(num, _)| (athlete_age_num - num).abs())
        .map(|(_, age_str)| age_str.clone());
    
    closest
}

fn time_to_seconds(value: &Data) -> Option<f64> {
    match value {
        Data::Float(f) => Some(*f),
        Data::Int(i) => Some(*i as f64),
        Data::DateTime(dt) => {
            // Excel stores time as fraction of a day
            // Convert to seconds: fraction_of_day * 24 hours * 60 minutes * 60 seconds
            let seconds = dt.as_f64() * 86400.0; // 24 * 60 * 60 = 86400 seconds per day
            Some(seconds)
        }
        Data::String(s) => {
            let s = s.trim();
            if s.is_empty() || s.eq_ignore_ascii_case("nan") {
                return None;
            }
            
            if s.contains(':') {
                let parts: Vec<&str> = s.split(':').collect();
                if parts.len() == 2 {
                    let minutes = parts[0].parse::<f64>().ok()?;
                    let seconds = parts[1].parse::<f64>().ok()?;
                    return Some(minutes * 60.0 + seconds);
                }
            }
            
            s.parse::<f64>().ok()
        }
        _ => None,
    }
}

fn parse_meet_file(file_path: &Path) -> Result<Vec<MeetResult>, Box<dyn std::error::Error>> {
    let filename = file_path.file_name()
        .and_then(|n| n.to_str())
        .ok_or("Invalid filename")?;
    
    let filename_clean = filename
        .replace(".xlsx", "")
        .replace(".xls", "");
    let parts: Vec<&str> = filename_clean.split('_').collect();
    
    if parts.len() < 5 {
        return Err(format!("Cannot parse filename: {}", filename).into());
    }
    
    let course = parts[2].to_string();
    let sex = parts[3].to_string();
    
    // Parse age range (format: XX-YY where YY is the age we want)
    let age_range = parts[4];
    let age_parts: Vec<&str> = age_range.split('-').collect();
    if age_parts.len() != 2 {
        return Err(format!("Invalid age range format: {}", age_range).into());
    }
    let age = age_parts[1].to_string(); // Get the YY part (e.g., "12" from "00-12")
    
    println!("  Parsing file: {} -> Sex: {}, Age: {}, Course: {}", filename, sex, age, course);
    
    let mut workbook: Xlsx<_> = open_workbook(file_path)?;
    let sheet_names: Vec<String> = workbook.sheet_names().iter().map(|s| s.to_string()).collect();
    
    let mut results = Vec::new();
    let mut results_count = 0;
    
    for sheet_name in &sheet_names {
        let event = match normalize_event_name(sheet_name) {
            Some(e) => e,
            None => continue,
        };
        
        if let Ok(range) = workbook.worksheet_range(sheet_name) {
            for row in range.rows() {
                if row.len() <= 9 {
                    continue;
                }
                
                // Column J (index 9) for times
                let time_seconds = match time_to_seconds(&row[9]) {
                    Some(t) if t > 0.0 => t,
                    _ => continue,
                };
                
                // Column E (index 4) for names
                let name = if row.len() > 4 {
                    match &row[4] {
                        Data::String(s) if !s.trim().is_empty() => s.trim().to_string(),
                        _ => String::new(),
                    }
                } else {
                    String::new()
                };
                
                results.push(MeetResult {
                    course: course.clone(),
                    sex: sex.clone(),
                    age: age.clone(),
                    event: event.clone(),
                    time: time_seconds,
                    name: name.clone(),
                });
                results_count += 1;
            }
        }
    }
    
    println!("    -> Found {} results", results_count);
    
    Ok(results)
}

fn load_time_standards(standards_file: &Path) -> Result<(HashMap<String, EventStandards>, HashMap<String, Vec<String>>), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(standards_file)?;
    let mut all_standards: HashMap<String, EventStandards> = HashMap::new();
    let mut event_orders: HashMap<String, Vec<String>> = HashMap::new();
    
    // Process both Mens and Womens tabs
    for gender in &["Mens", "Womens"] {
        let mut standards: EventStandards = HashMap::new();
        let mut event_order: Vec<String> = Vec::new();
        
        if let Ok(range) = workbook.worksheet_range(gender) {
            let mut age_groups: Vec<String> = Vec::new();
            
            // Read header row to get age groups (columns B onwards)
            if let Some(header_row) = range.rows().next() {
                println!("\nDEBUG: Processing {} tab", gender);
                println!("  Header row cells:");
                for (idx, cell) in header_row.iter().enumerate() {
                    let cell_str = match cell {
                        Data::String(s) => s.clone(),
                        Data::Int(i) => i.to_string(),
                        Data::Float(f) => f.to_string(),
                        Data::Empty => "(empty)".to_string(),
                        _ => format!("{:?}", cell),
                    };
                    println!("    Column {}: '{}'", idx, cell_str);
                }
                
                for cell in header_row.iter().skip(1) {
                    let age_str = match cell {
                        Data::String(s) => s.trim().to_string(),
                        Data::Int(i) => i.to_string(),
                        Data::Float(f) => f.to_string(),
                        _ => String::new(),
                    };
                    
                    if !age_str.is_empty() {
                        age_groups.push(normalize_age(&age_str));
                    }
                }
            }
            
            println!("  Age groups found: {:?}", age_groups);
            
            // Process data rows
            let mut row_count = 0;
            for row in range.rows().skip(1) {
                if row.is_empty() {
                    continue;
                }
                
                // Column A - Event name
                let event_str = match &row[0] {
                    Data::String(s) => s.trim(),
                    _ => continue,
                };
                
                if event_str.is_empty() {
                    continue;
                }
                
                let normalized_event = match normalize_event_name(event_str) {
                    Some(e) => e,
                    None => continue,
                };
                
                if row_count < 3 {
                    println!("  Sample event row {}: '{}' -> '{}'", row_count, event_str, normalized_event);
                }
                
                event_order.push(normalized_event.clone());
                
                // Read times for each age group (columns B onwards)
                let mut age_standards: AgeGroupStandards = HashMap::new();
                
                for (idx, age_group) in age_groups.iter().enumerate() {
                    let col_idx = idx + 1; // Skip event column
                    if col_idx < row.len() {
                        let cell_value = &row[col_idx];
                        if row_count < 1 && idx < 3 {
                            println!("    Age '{}' (col {}): cell = {:?}", age_group, col_idx, cell_value);
                        }
                        
                        if let Some(time_value) = time_to_seconds(cell_value) {
                            age_standards.insert(age_group.clone(), time_value);
                            if row_count < 1 && idx < 3 {
                                println!("      -> Parsed as {:.2}s", time_value);
                            }
                        } else if row_count < 1 && idx < 3 {
                            println!("      -> Failed to parse");
                        }
                    }
                }
                
                if row_count < 1 {
                    println!("    Total ages with standards for this event: {}", age_standards.len());
                }
                
                standards.insert(normalized_event, age_standards);
                row_count += 1;
            }
            
            println!("  Total events loaded: {}", row_count);
        }
        
        let gender_key = if *gender == "Mens" { "Men" } else { "Women" };
        all_standards.insert(gender_key.to_string(), standards);
        event_orders.insert(gender_key.to_string(), event_order);
    }
    
    Ok((all_standards, event_orders))
}

fn count_qualifiers(
    meet_results: &[MeetResult],
    standards: &HashMap<String, EventStandards>,
) -> HashMap<StandardKey, usize> {
    let mut qualifier_counts: HashMap<StandardKey, usize> = HashMap::new();
    let mut matches_found = 0;
    let mut no_standard_count = 0;
    
    for result in meet_results {
        // Get standards for this gender
        if let Some(gender_standards) = standards.get(&result.sex) {
            // Get standards for this event
            if let Some(event_standards) = gender_standards.get(&result.event) {
                // Check if there's a qualifying time for this age
                if let Some(&qualifying_time) = event_standards.get(&result.age) {
                    if result.time <= qualifying_time {
                        let key = (result.sex.clone(), result.age.clone(), result.event.clone());
                        *qualifier_counts.entry(key).or_insert(0) += 1;
                        matches_found += 1;
                    }
                } else {
                    no_standard_count += 1;
                }
            }
        }
    }
    
    println!("DEBUG: Found {} qualifying times", matches_found);
    println!("DEBUG: {} results had no matching standard", no_standard_count);
    
    qualifier_counts
}

fn count_unique_qualifiers(
    meet_results: &[MeetResult],
    standards: &HashMap<String, EventStandards>,
) -> HashMap<(String, String), HashSet<String>> {
    let mut unique_qualifiers: HashMap<(String, String), HashSet<String>> = HashMap::new();
    
    for result in meet_results {
        if result.name.is_empty() {
            continue;
        }
        
        if let Some(gender_standards) = standards.get(&result.sex) {
            if let Some(event_standards) = gender_standards.get(&result.event) {
                let available_ages: Vec<String> = event_standards.keys().cloned().collect();
                
                // Find best matching age
                if let Some(matched_age) = find_best_age_match(&result.age, &available_ages) {
                    if let Some(&qualifying_time) = event_standards.get(&matched_age) {
                        if result.time <= qualifying_time {
                            // Use the MATCHED age, not the original age
                            let key = (result.sex.clone(), matched_age.clone());
                            unique_qualifiers.entry(key)
                                .or_insert_with(HashSet::new)
                                .insert(result.name.clone());
                        }
                    }
                }
            }
        }
    }
    
    unique_qualifiers
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Show current directory for debugging
    let current_dir = std::env::current_dir()?;
    println!("Running from: {:?}", current_dir);
    
    // Check if standards file exists
    let standards_path = Path::new(TIME_STANDARDS_FILE);
    let full_path = current_dir.join(standards_path);
    println!("Looking for standards file at: {:?}", full_path);
    
    if !standards_path.exists() {
        // List files in current directory to help debug
        println!("\nFiles in current directory:");
        if let Ok(entries) = std::fs::read_dir(&current_dir) {
            for entry in entries.flatten() {
                if let Ok(file_type) = entry.file_type() {
                    let prefix = if file_type.is_dir() { "[DIR] " } else { "" };
                    println!("  {}{}", prefix, entry.file_name().to_string_lossy());
                }
            }
        }
        return Err(format!("Time standards file not found: {}", TIME_STANDARDS_FILE).into());
    }
    
    println!("Loading time standards from {}...", TIME_STANDARDS_FILE);
    let (standards, event_orders) = load_time_standards(Path::new(TIME_STANDARDS_FILE))?;
    
    for (gender, gender_standards) in &standards {
        println!("Loaded {} events for {}", gender_standards.len(), gender);
        
        // Show what ages are in the standards
        let mut std_ages: HashSet<String> = HashSet::new();
        for event_standards in gender_standards.values() {
            for age in event_standards.keys() {
                std_ages.insert(age.clone());
            }
        }
        let mut std_ages_vec: Vec<_> = std_ages.into_iter().collect();
        std_ages_vec.sort_by_key(|a| a.parse::<i32>().unwrap_or(999));
        println!("  Ages in standards: {:?}", std_ages_vec);
        
        // Show sample events
        let sample_events: Vec<_> = gender_standards.keys().take(5).collect();
        println!("  Sample events: {:?}", sample_events);
    }
    
    println!("\nSearching for meet files in {}...", DATA_FOLDER);
    
    // Check if data folder exists
    if !Path::new(DATA_FOLDER).exists() {
        return Err(format!("Data folder not found: {}", DATA_FOLDER).into());
    }
    
    let mut meet_files = Vec::new();
    
    for entry in std::fs::read_dir(DATA_FOLDER)? {
        let entry = entry?;
        let path = entry.path();
        if let Some(filename) = path.file_name().and_then(|n| n.to_str()) {
            if filename.starts_with("CAN-MBSK_") && 
               (filename.ends_with(".xlsx") || filename.ends_with(".xls")) {
                meet_files.push(path);
            }
        }
    }
    
    println!("Found {} meet files", meet_files.len());
    
    if meet_files.is_empty() {
        return Err("No meet files found!".into());
    }
    
    println!("\nParsing meet files...");
    let mut all_results = Vec::new();
    
    for file_path in &meet_files {
        println!("  Processing {:?}...", file_path.file_name());
        match parse_meet_file(file_path) {
            Ok(results) => {
                all_results.extend(results);
            }
            Err(e) => println!("  Error: {}", e),
        }
    }
    
    println!("\nTotal results extracted: {}", all_results.len());
    
    // Debug: Show sample of what we parsed
    if !all_results.is_empty() {
        println!("\nSample results:");
        for result in all_results.iter().take(3) {
            println!("  Sex: {}, Age: {}, Event: {}, Time: {:.2}s", 
                     result.sex, result.age, result.event, result.time);
        }
    }
    
    // Debug: Show what ages and events we have
    let mut ages: HashSet<String> = all_results.iter().map(|r| r.age.clone()).collect();
    let mut ages_vec: Vec<_> = ages.iter().cloned().collect();
    ages_vec.sort_by_key(|a| a.parse::<i32>().unwrap_or(999));
    println!("\nAges found in meet data: {:?}", ages_vec);
    
    let events: HashSet<String> = all_results.iter().map(|r| r.event.clone()).collect();
    println!("Events found in meet data: {:?}", events.iter().take(5).collect::<Vec<_>>());
    
    println!("\nCounting qualifiers...");
    let qualifier_counts = count_qualifiers(&all_results, &standards);
    let unique_qualifiers = count_unique_qualifiers(&all_results, &standards);
    
    println!("Found {} qualifier count entries", qualifier_counts.len());
    
    // Count total unique athletes per gender/age (using matched ages)
    let mut total_athletes: HashMap<(String, String), HashSet<String>> = HashMap::new();
    for result in &all_results {
        if !result.name.is_empty() {
            // Find best matching age for this result
            if let Some(gender_standards) = standards.get(&result.sex) {
                // Get any event to find available ages
                if let Some((_, event_standards)) = gender_standards.iter().next() {
                    let available_ages: Vec<String> = event_standards.keys().cloned().collect();
                    if let Some(matched_age) = find_best_age_match(&result.age, &available_ages) {
                        let key = (result.sex.clone(), matched_age);
                        total_athletes.entry(key)
                            .or_insert_with(HashSet::new)
                            .insert(result.name.clone());
                    }
                }
            }
        }
    }
    
    // Create output workbook
    let mut workbook = Workbook::new();
    
    // Process each gender
    for gender in &["Men", "Women"] {
        let sheet_name = if *gender == "Men" { "Mens" } else { "Womens" };
        let sheet = workbook.add_worksheet();
        sheet.set_name(sheet_name)?;
        
        // Get standards and event order for this gender
        let gender_standards = match standards.get(*gender) {
            Some(s) => s,
            None => continue,
        };
        
        let event_order = match event_orders.get(*gender) {
            Some(o) => o,
            None => continue,
        };
        
        // Collect all age groups
        let mut age_groups: HashSet<String> = HashSet::new();
        for event_standards in gender_standards.values() {
            for age in event_standards.keys() {
                age_groups.insert(age.clone());
            }
        }
        
        let mut age_groups_vec: Vec<String> = age_groups.into_iter().collect();
        age_groups_vec.sort_by_key(|a| a.parse::<i32>().unwrap_or(999));
        
        // Write headers
        sheet.write_string(0, 0, "Event")?;
        for (i, age) in age_groups_vec.iter().enumerate() {
            sheet.write_string(0, (i + 1) as u16, age)?;
        }
        
        // Write data rows following event order
        let mut row = 1u32;
        for event in event_order {
            sheet.write_string(row, 0, event)?;
            
            for (col, age) in age_groups_vec.iter().enumerate() {
                let key = (gender.to_string(), age.clone(), event.clone());
                let count = qualifier_counts.get(&key).copied().unwrap_or(0);
                sheet.write_number(row, (col + 1) as u16, count as f64)?;
            }
            
            row += 1;
        }
        
        // Add summary rows
        row += 1;
        sheet.write_string(row, 0, "Total Unique Athletes")?;
        for (col, age) in age_groups_vec.iter().enumerate() {
            let key = (gender.to_string(), age.clone());
            let count = total_athletes.get(&key).map(|s| s.len()).unwrap_or(0);
            sheet.write_number(row, (col + 1) as u16, count as f64)?;
        }
        
        row += 1;
        sheet.write_string(row, 0, "Unique Qualifiers")?;
        for (col, age) in age_groups_vec.iter().enumerate() {
            let key = (gender.to_string(), age.clone());
            let count = unique_qualifiers.get(&key).map(|s| s.len()).unwrap_or(0);
            sheet.write_number(row, (col + 1) as u16, count as f64)?;
        }
    }
    
    workbook.save(OUTPUT_FILE)?;
    println!("\nAnalysis complete! Results saved to {}", OUTPUT_FILE);
    
    Ok(())
}