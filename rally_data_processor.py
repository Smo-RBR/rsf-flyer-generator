import csv
import re

def parse_time_to_seconds(time_str):
    if not time_str:
        return None
    if ':' in time_str:
        parts = time_str.split(':')
        if len(parts) == 2:
            minutes = float(parts[0])
            seconds = float(parts[1])
            return minutes * 60 + seconds
        elif len(parts) == 3: # Handle hh:mm:ss.sss if it ever appears
            hours = float(parts[0])
            minutes = float(parts[1])
            seconds = float(parts[2])
            return hours * 3600 + minutes * 60 + seconds
    return float(time_str) # Assume it's already in seconds if no colon

def format_seconds_to_mmss(seconds):
    if seconds is None:
        return ""
    minutes = int(seconds // 60)
    remaining_seconds = seconds % 60
    return f"{minutes:02d}:{remaining_seconds:06.3f}"

def normalize_name_casing(name):
    if not name:
        return name
    return name[0].upper() + name[1:]

def load_csv(file_content):
    reader = csv.DictReader(file_content.splitlines(), delimiter=';')
    data = []
    for row in reader:
        processed_row = {k.strip(): v.strip() for k, v in row.items()}
        # Normalize casing for User name and Real name upon loading
        if 'User name' in processed_row:
            processed_row['User name'] = normalize_name_casing(processed_row['User name'])
        if 'Real name' in processed_row:
            processed_row['Real name'] = normalize_name_casing(processed_row['Real name'])
        if 'user_name' in processed_row:
            processed_row['user_name'] = normalize_name_casing(processed_row['user_name'])
        if 'real_name' in processed_row:
            processed_row['real_name'] = normalize_name_casing(processed_row['real_name'])
        data.append(processed_row)
    return data

def validate_stages_data(stages_data):
    errors = []
    seen_stages = {} # To check for sequential SS and duplicates

    for i, row in enumerate(stages_data):
        row_num = i + 2 # Account for header row

        # Check for essential columns, allowing time3 and Progress to be empty for retirements
        required_cols_always_present = ['SS', 'Stage name', 'User name']
        for col in required_cols_always_present:
            if col not in row or not row[col]:
                errors.append(f"Stages file, Row {row_num}: Missing or empty value in column '{col}'.")

        # Validate SS number
        try:
            ss_num = int(row['SS'])
            if ss_num <= 0:
                errors.append(f"Stages file, Row {row_num}: 'SS' must be a positive integer.")
            
            user_stage_key = (row.get('User name'), ss_num)
            if user_stage_key in seen_stages:
                errors.append(f"Stages file, Row {row_num}: Duplicate entry for User '{row.get('User name')}' on SS {ss_num}.")
            seen_stages[user_stage_key] = True

        except (ValueError, TypeError):
            errors.append(f"Stages file, Row {row_num}: 'SS' is not a valid integer.")

        # Validate time3 and Progress based on retirement rules
        time3_val = row.get('time3')
        progress_val = row.get('Progress')

        if progress_val == 'F': # Driver finished stage
            if not time3_val:
                errors.append(f"Stages file, Row {row_num}: 'Progress' is 'F' but 'time3' is missing. Inconsistent finish status.")
            else:
                try:
                    parse_time_to_seconds(time3_val)
                except ValueError:
                    errors.append(f"Stages file, Row {row_num}: 'time3' has an invalid format ('{time3_val}'). Expected mm:ss.sss or seconds.")
        elif progress_val == '': # Driver retired in this stage
            if time3_val:
                errors.append(f"Stages file, Row {row_num}: 'Progress' is empty but 'time3' is present. Inconsistent retirement status.")
            # No need to check time1/time2/time3 for presence here, as per rules they can be empty for retirements.
        else:
            errors.append(f"Stages file, Row {row_num}: 'Progress' has an invalid value ('{progress_val}'). Expected 'F' or empty.")

    return errors

def validate_final_data(final_data):
    errors = []
    seen_ranks = {} # To check for sequential # and duplicates
    seen_users = set() # To check for duplicate users

    for i, row in enumerate(final_data):
        row_num = i + 2 # Account for header row

        # Check for essential columns
        required_cols = ['#', 'user_name', 'real_name', 'time3']
        for col in required_cols:
            if col not in row or not row[col]:
                errors.append(f"Final file, Row {row_num}: Missing or empty value in column '{col}'.")

        # Validate rank
        try:
            rank = int(row['#'])
            if rank <= 0:
                errors.append(f"Final file, Row {row_num}: '#' must be a positive integer.")
            if rank in seen_ranks:
                errors.append(f"Final file, Row {row_num}: Duplicate rank '{rank}'.")
            seen_ranks[rank] = True
        except (ValueError, TypeError):
            errors.append(f"Final file, Row {row_num}: '#' is not a valid integer.")

        # Validate user_name uniqueness
        user_name = row.get('user_name')
        if user_name:
            if user_name in seen_users:
                errors.append(f"Final file, Row {row_num}: Duplicate 'user_name' ('{user_name}').")
            seen_users.add(user_name)

        # Validate time3 format
        if 'time3' in row and row['time3']:
            try:
                # Attempt to parse time3 to seconds to ensure it's a valid time
                parse_time_to_seconds(row['time3'])
            except ValueError:
                errors.append(f"Final file, Row {row_num}: 'time3' has an invalid format ('{row['time3']}'). Expected mm:ss.sss or seconds.")
    
    # Check for gaps in ranks
    if seen_ranks:
        max_rank = max(seen_ranks.keys())
        for r in range(1, max_rank + 1):
            if r not in seen_ranks:
                errors.append(f"Final file: Gap in ranks, rank {r} is missing.")

    return errors

def cross_validate_data(stages_data, final_data):
    errors = []
    
    # Get all unique drivers from stages data
    stages_drivers = set()
    for row in stages_data:
        user_name = row.get('User name')
        real_name = row.get('Real name')
        if user_name and real_name:
            stages_drivers.add((user_name, real_name))
    
    # Get all unique drivers from final data
    final_drivers = set()
    for row in final_data:
        user_name = row.get('user_name')
        real_name = row.get('real_name')
        if user_name and real_name:
            final_drivers.add((user_name, real_name))

    # Check if all drivers in final are also in stages
    for user_name, real_name in final_drivers:
        if (user_name, real_name) not in stages_drivers:
            errors.append(f"Cross-validation: Driver '{user_name}' ('{real_name}') found in final results but not in stages data.")
    
    # Check if all drivers in stages that finished are in final (those with 'F' in last stage)
    # This is more complex as a driver might retire before the final stage.
    # For now, just check if names match.

    # Casing is now handled during load_csv, so no need for explicit casing checks here.
    return errors

import sys

if __name__ == "__main__":
    stages_file_path = "../../Downloads/DE-DCR-69-stages.csv"
    final_file_path = "../../Downloads/DE-DCR-69-final.csv"

    try:
        with open(stages_file_path, 'r', encoding='utf-8') as f:
            stages_content = f.read()
        with open(final_file_path, 'r', encoding='utf-8') as f:
            final_content = f.read()
    except FileNotFoundError as e:
        print(f"Error: File not found - {e.filename}")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading files: {e}")
        sys.exit(1)
    
    stages_data = load_csv(stages_content)
    final_data = load_csv(final_content)

    stages_errors = validate_stages_data(stages_data)
    final_errors = validate_final_data(final_data)
    cross_errors = cross_validate_data(stages_data, final_data)

    all_errors = stages_errors + final_errors + cross_errors

    if all_errors:
        print("Data validation failed with the following errors:")
        for error in all_errors:
            print(f"- {error}")
        sys.exit(1)
    else:
        print("Data validation successful. No errors found.")
        
        # Proceed with report generation
        report_style = "Sporty" # User selected style

        # Prepare data for report generation
        stages_by_ss = {}
        driver_overall_times = {} # To track total time for each driver
        driver_mentions = {} # To track mentions for each driver

        # Initialize driver_overall_times and driver_mentions from final_data
        for driver_final in final_data:
            user_name = driver_final['user_name']
            driver_overall_times[user_name] = parse_time_to_seconds(driver_final['time3'])
            driver_mentions[user_name] = 0

        # Group stages data by SS and collect driver info
        for row in stages_data:
            ss_num = int(row['SS'])
            if ss_num not in stages_by_ss:
                stages_by_ss[ss_num] = []
            stages_by_ss[ss_num].append(row)
            
            # Ensure all drivers from stages are in driver_mentions, even if they retired early
            user_name = row['User name']
            if user_name not in driver_mentions:
                driver_mentions[user_name] = 0

        # Sort stages by SS number
        sorted_ss_nums = sorted(stages_by_ss.keys())

        full_report = []

        # Generate report for each stage
        for ss_num in sorted_ss_nums:
            stage_name = stages_by_ss[ss_num][0]['Stage name']
            stage_results = stages_by_ss[ss_num]
            
            def sort_key(entry):
                time3 = parse_time_to_seconds(entry.get('time3'))
                return time3 if time3 is not None else float('inf')

            stage_results_sorted = sorted(stage_results, key=sort_key)

            stage_report_lines = []
            stage_report_lines.append(f"## Etappe {ss_num}: {stage_name} 🏁")

            # Identify stage winner
            winner = None
            for driver_result in stage_results_sorted:
                if driver_result.get('Progress') == 'F' and driver_result.get('time3'):
                    winner = driver_result
                    break
            
            if winner:
                winner_name = normalize_name_casing(winner['User name'])
                winner_time = format_seconds_to_mmss(parse_time_to_seconds(winner['time3']))
                stage_report_lines.append(f"Die **Etappe {ss_num}** auf **{stage_name}** startete mit einem Adrenalinkick! **{winner_name}** zeigte eine Meisterleistung und sicherte sich mit einer phänomenalen Zeit von **{winner_time}** den Etappensieg!")
                driver_mentions[winner_name] += 1

            # Identify duels and close finishes
            for i in range(len(stage_results_sorted) - 1):
                driver1 = stage_results_sorted[i]
                driver2 = stage_results_sorted[i+1]

                time1_s = parse_time_to_seconds(driver1.get('time3'))
                time2_s = parse_time_to_seconds(driver2.get('time3'))

                if time1_s is not None and time2_s is not None:
                    diff = abs(time1_s - time2_s)
                    if diff < 5.0: # Example threshold for a close duel (5 seconds)
                        driver1_name = normalize_name_casing(driver1['User name'])
                        driver2_name = normalize_name_casing(driver2['User name'])
                        stage_report_lines.append(f"Ein packendes Duell entbrannte zwischen **{driver1_name}** und **{driver2_name}**! Sie lieferten sich einen Kampf auf Messers Schneide, getrennt durch hauchdünne **{format_seconds_to_mmss(diff)}** Sekunden!")
                        driver_mentions[driver1_name] += 1
                        driver_mentions[driver2_name] += 1
            
            # Identify retirements and comments
            for driver_result in stage_results_sorted:
                user_name = normalize_name_casing(driver_result['User name'])
                progress = driver_result.get('Progress')
                comment = driver_result.get('Comment')

                if progress == '': # Driver retired
                    retirement_reason = ""
                    if not driver_result.get('time1'):
                        retirement_reason = "bereits vor dem ersten Zwischenzeitpunkt"
                    elif not driver_result.get('time2'):
                        retirement_reason = "zwischen dem ersten und zweiten Zwischenzeitpunkt"
                    elif not driver_result.get('time3'):
                        retirement_reason = "im letzten Drittel der Etappe"
                    
                    comment_insight = ""
                    if comment:
                        comment_insight = f" Ihr Kommentar: *'{comment.strip()}'* sprach Bände über die Herausforderung."
                    
                    stage_report_lines.append(f"Ein bitteres Aus für **{user_name}**! Der Fahrer musste {retirement_reason} auf dieser gnadenlosen Etappe aufgeben.{comment_insight}")
                    driver_mentions[user_name] += 1
                
                # Penalties
                penalty = driver_result.get('Penalty')
                if penalty and float(penalty) > 0:
                    user_name = normalize_name_casing(driver_result['User name'])
                    stage_report_lines.append(f"Ein herber Dämpfer für **{user_name}**, der eine **{penalty}**-Sekunden-Strafe kassierte! Jeder Wimpernschlag zählt in dieser Rallye!")
                    driver_mentions[user_name] += 1

            # Add general commentary to reach 8-10 sentences if needed
            while len(stage_report_lines) < 8:
                stage_report_lines.append("Die Piloten meisterten das anspruchsvolle Terrain mit Bravour und trieben ihre Boliden bis an die Grenzen des Machbaren.")
            
            full_report.extend(stage_report_lines)
            full_report.append("\n") # Add a blank line between stages for Discord

        # Final summary
        full_report.append("## Endstand der Rallye 🏆")
        
        # Sort final data by rank
        final_data_sorted = sorted(final_data, key=lambda x: int(x['#']))

        winner_final = final_data_sorted[0]
        winner_name_final = normalize_name_casing(winner_final['user_name'])
        winner_time_final = format_seconds_to_mmss(parse_time_to_seconds(winner_final['time3']))
        full_report.append(f"Nach einer kräftezehrenden Rallye steht der Champion fest! Ein riesiger Applaus für **{winner_name_final}**, der mit einer atemberaubenden Gesamtzeit von **{winner_time_final}** den Gesamtsieg einfuhr!")
        driver_mentions[winner_name_final] += 1

        # Mention top performers
        if len(final_data_sorted) > 1:
            second_place = final_data_sorted[1]
            second_name = normalize_name_casing(second_place['user_name'])
            second_time = format_seconds_to_mmss(parse_time_to_seconds(second_place['time3']))
            full_report.append(f"**{second_name}** zeigte eine beeindruckende Leistung und sicherte sich mit **{second_time}** einen verdienten zweiten Platz. Ihre Konstanz war bemerkenswert!")
            driver_mentions[second_name] += 1
        
        if len(final_data_sorted) > 2:
            third_place = final_data_sorted[2]
            third_name = normalize_name_casing(third_place['user_name'])
            third_time = format_seconds_to_mmss(parse_time_to_seconds(third_place['time3']))
            full_report.append(f"Das Podium komplettiert **{third_name}**, der mit unglaublicher Zähigkeit und einer Zeit von **{third_time}** den dritten Rang eroberte. Eine starke Vorstellung!")
            driver_mentions[third_name] += 1

        # Check for drivers not mentioned twice and add general mentions
        for driver, mentions in driver_mentions.items():
            if mentions < 2:
                full_report.append(f"Auch **{driver}** trug mit seinem Einsatz und seiner Entschlossenheit maßgeblich zur Spannung dieser Rallye bei.")
        
        full_report.append("\nDiese Rallye war ein ultimativer Härtetest für Können, Ausdauer und Nerven. Von packenden Kopf-an-Kopf-Rennen bis zu dramatischen Ausfällen – sie lieferte Action pur und unvergessliche Momente. Herzlichen Glückwunsch an alle Teilnehmer für ihre herausragenden Leistungen! 🎉")

        # Join with double newline for better Discord paragraph separation
        print("\n\n".join(full_report))
