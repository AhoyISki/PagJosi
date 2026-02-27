#![windows_subsystem = "windows"]
use std::{
    collections::HashMap,
    num::ParseFloatError,
    path::{Path, PathBuf},
    sync::{Arc, LazyLock, Mutex},
};

use csv::{StringRecord, StringRecordsIntoIter};
use iced::{
    Alignment, Border, Length, Padding, Task, Theme,
    border::Radius,
    widget::{Column, Container, button, column, combo_box, container, row, text, text_input},
};
use regex::Regex;
use rfd::{AsyncFileDialog, FileHandle};
use umya_spreadsheet::{
    BorderStyleValues, Color, Font, Spreadsheet, Style, VerticalAlignmentValues,
};
use unaccent::unaccent;

macro_rules! main_container {
    ($contained:expr) => {
        Container::new($contained)
            .padding(Padding::new(10.0))
            .style(|theme| {
                let palette = Theme::extended_palette(theme);
                container::Style::default()
                    .background(palette.background.weakest.color)
                    .border(Border {
                        color: palette.background.weakest.color,
                        width: 0.0,
                        radius: Radius::new(10.0),
                    })
            })
    };
}

fn main() {
    _ = iced::application(Application::default, Application::update, Application::view)
        .theme(Theme::CatppuccinLatte)
        .run();
}

#[derive(Default)]
struct Application {
    files: HashMap<&'static str, PathBuf>,
    is_picking_file: bool,
    result: Option<Result<(), Arc<dyn std::error::Error + Send + Sync>>>,
    directory: Option<PathBuf>,
    assoc_request: Option<AssocRequest>,
    began_processing: bool,
}

impl Application {
    fn update(&mut self, message: Message) -> Task<Message> {
        let handle_result = |result| match result {
            Ok(Some((parts, reason))) => {
                Message::SheetAssocRequested(Arc::new(Mutex::new(Some(parts))), reason)
            }
            result @ (Ok(None) | Err(_)) => Message::Result(result.map(|_| ()).map_err(Arc::from)),
        };

        match message {
            Message::PickFile(name, ext) if !self.is_picking_file => {
                self.is_picking_file = true;
                let pick_file = {
                    let dialog = AsyncFileDialog::new().add_filter(ext, &[ext]);
                    if let Some(path) = self.directory.as_ref() {
                        dialog.set_directory(path).pick_file()
                    } else {
                        dialog.pick_file()
                    }
                };

                Task::perform(pick_file, move |file| Message::FilePicked(name, file))
            }
            Message::PickFile(..) => Task::none(),
            Message::FilePicked(name, file_handle) => {
                if let Some(file_handle) = file_handle {
                    self.files.insert(name, file_handle.path().to_path_buf());
                    if let Some(parent) = file_handle.path().parent() {
                        self.directory = Some(parent.to_path_buf());
                    }
                }
                self.is_picking_file = false;
                Task::none()
            }
            Message::StartRepassAddition => {
                self.began_processing = true;

                let parts = match ProcessParts::new(
                    &self.files["Vendas"],
                    &self.files["Financeiro"],
                    &self.files["Repasses"],
                ) {
                    Ok(parts) => parts,
                    Err(err) => return Task::done(Message::Result(Err(Arc::from(err)))),
                };

                Task::perform(process_csvs(parts), handle_result)
            }
            Message::Result(result) => {
                self.result = Some(result);
                Task::none()
            }
            Message::SheetAssocRequested(parts, reason) => {
                let associations = ASSOCIATIONS.lock().unwrap();

                let parts = parts.lock().unwrap().take().unwrap();
                let condo_name = match &reason {
                    AssocRequestReason::NotFound(condo) => condo.clone(),
                    AssocRequestReason::NewCondo(obs) => match obs.split_once("-") {
                        Some((before, _)) => before.trim().to_string(),
                        None => obs.clone(),
                    },
                };

                let options = parts
                    .spreadsheet
                    .get_sheet_collection_no_check()
                    .iter()
                    .map(|sheet| sheet.get_name().to_string())
                    .collect();

                let combo_state = combo_box::State::with_selection(
                    options,
                    associations
                        .get(&condo_name)
                        .and_then(|assoc| assoc.as_ref().map(|(sheet, _)| sheet)),
                );

                self.assoc_request = Some(AssocRequest {
                    parts,
                    reason,
                    combo_state,
                    assoc_sheet: None,
                    condo_name,
                });

                Task::none()
            }
            Message::ConfirmSheetAssoc(obs, assoc_sheet, condo_name) => {
                let mut associations = ASSOCIATIONS.lock().unwrap();
                let mut request = self.assoc_request.take().unwrap();

                // "mut" to prevent a stupid panic.
                if request
                    .parts
                    .spreadsheet
                    .get_sheet_by_name_mut(&assoc_sheet)
                    .is_none()
                    && let Err(err) = request.parts.spreadsheet.new_sheet(&assoc_sheet)
                {
                    return Task::done(Message::Result(Err(Arc::from(Box::from(err)))));
                }

                if let Some(obs) = obs {
                    associations.insert(obs, Some((assoc_sheet.clone(), condo_name.clone())));
                }
                associations.insert(condo_name.clone(), Some((assoc_sheet, condo_name)));

                Task::perform(process_csvs(request.parts), handle_result)
            }
            Message::SkipAssoc(skipped) => {
                let mut associations = ASSOCIATIONS.lock().unwrap();
                associations.insert(skipped, None);

                let request = self.assoc_request.take().unwrap();
                Task::perform(process_csvs(request.parts), handle_result)
            }
            Message::SetAssocSheet(assoc_sheet) => {
                let mut request = self.assoc_request.take().unwrap();

                let mut options = request.combo_state.into_options();

                if let Some(old) = request.assoc_sheet.as_ref() {
                    options.retain(|item| item != old);
                };

                options.insert(0, assoc_sheet.clone());

                request.combo_state = combo_box::State::with_selection(options, Some(&assoc_sheet));
                request.assoc_sheet = Some(assoc_sheet);

                self.assoc_request = Some(request);
                Task::none()
            }
            Message::SetCondoName(condo_name) => {
                let request = self.assoc_request.as_mut().unwrap();
                request.condo_name = condo_name;
                Task::none()
            }
        }
    }

    fn file_selector(&self, name: &'static str, ext: &'static str) -> Container<'_, Message> {
        let title = text(name).size(20);

        let chosen_file = {
            let chosen_file = if let Some(path) = self.files.get(name) {
                text(path.to_string_lossy())
            } else {
                text("Escolha um arquivo...")
            };

            chosen_file
                .align_x(Alignment::Center)
                .align_y(Alignment::Center)
                .height(32.0)
                .width(Length::Fill)
        };

        let picker = Container::new(
            button(row![
                chosen_file,
                text("📂").align_y(Alignment::Center).height(Length::Fill)
            ])
            .style(button_style())
            .height(32.0)
            .on_press_maybe((!self.began_processing).then_some(Message::PickFile(name, ext))),
        );

        Container::new(column![title, picker])
            .padding(Padding::new(10.0))
            .style(|theme| {
                let palette = Theme::extended_palette(theme);
                container::Style::default()
                    .background(palette.background.weak.color)
                    .border(Border {
                        color: palette.background.weak.color,
                        width: 0.0,
                        radius: Radius::new(10.0),
                    })
            })
    }

    fn view(&self) -> Column<'_, Message> {
        let title = Container::new(text("PagJosi").size(32)).padding(Padding::new(10.0));

        let confirm_button = (!self.began_processing).then(|| {
            Some(
                button("Adicionar novos repasses")
                    .style(button_style())
                    .height(32.0)
                    .on_press_maybe(
                        if self.files.contains_key("Vendas")
                            && self.files.contains_key("Financeiro")
                            && self.files.contains_key("Repasses")
                        {
                            Some(Message::StartRepassAddition)
                        } else {
                            None
                        },
                    ),
            )
        });

        let column = column![
            title,
            main_container!(
                column![
                    self.file_selector("Vendas", "csv"),
                    self.file_selector("Financeiro", "csv"),
                    self.file_selector("Repasses", "xlsx"),
                    confirm_button
                ]
                .spacing(10.0)
                .align_x(Alignment::End)
            )
        ]
        .spacing(10.0)
        .padding(10.0);

        match self.result.as_ref() {
            Some(Ok(())) => column
                .push(main_container!(
                    text("Descrição dos repasses completa!")
                        .align_x(Alignment::Center)
                        .align_y(Alignment::Center)
                        .height(32.0)
                        .width(Length::Fill)
                ))
                .width(Length::Fill),
            Some(Err(error)) => column
                .push(main_container!(
                    text(error.to_string())
                        .align_x(Alignment::Center)
                        .align_y(Alignment::Center)
                        .height(32.0)
                        .width(Length::Fill)
                ))
                .width(Length::Fill),
            None => match self.assoc_request.as_ref() {
                Some(request) => column.push({
                    main_container!(match &request.reason {
                        AssocRequestReason::NotFound(condo_name) => {
                            let condo_name = condo_name.clone();

                            column![
                                text(format!(
                                    "Não foi encontrada uma aba de repasses que bate com o \
                                     condomínio \"{condo_name}\". Escolha um nome de uma aba para \
                                     associar a esse condomínio.\nEste pode ser o nome de uma \
                                     nova aba. Se for o caso, esta será adicionada as abas no \
                                     arquivo de repasses."
                                ))
                                .width(Length::Fill),
                                combo_box(
                                    &request.combo_state,
                                    &format!(
                                        "Escolha um nome para uma aba nova ou já existente (e.g. \
                                         \"70 - {condo_name}\")",
                                    ),
                                    None,
                                    {
                                        let condo_name = condo_name.clone();
                                        move |assoc_sheet| {
                                            Message::ConfirmSheetAssoc(
                                                None,
                                                assoc_sheet,
                                                condo_name.clone(),
                                            )
                                        }
                                    }
                                )
                                .on_input(Message::SetAssocSheet),
                                button("Pular")
                                    .style(button_style())
                                    .height(32.0)
                                    .on_press(Message::SkipAssoc(condo_name)),
                            ]
                            .spacing(10.0)
                            .align_x(Alignment::End)
                        }
                        AssocRequestReason::NewCondo(obs) => {
                            let condo_name = request.condo_name.clone();
                            let obs = obs.clone();

                            column![
                                text(format!(
                                    "Foi encontrado um novo condomínio (descrição é \"NOVO \
                                     CONDOMINIO\") com a observação:\n\"{obs}\".\n Você deve \
                                     decidir um nome para a aba de repasses e qual será o nome \
                                     desse condomínio na tabela dentro desta aba.\nUm nome para \
                                     este condomínio já foi escolhido ({condo_name}), mas se \
                                     quiser, pode altera-lo se estiver incorreto. (e.g. trocar \
                                     \"TARRAF VIVA VOTUPORANGA\" por \"TARRAF VIVA\")"
                                ))
                                .width(Length::Fill),
                                text_input(&condo_name, &condo_name)
                                    .on_submit_maybe(request.assoc_sheet.as_ref().map(
                                        |assoc_sheet| Message::ConfirmSheetAssoc(
                                            Some(obs.clone()),
                                            assoc_sheet.clone(),
                                            condo_name.clone(),
                                        )
                                    ))
                                    .on_input(|input| Message::SetCondoName(
                                        input.trim().to_string()
                                    )),
                                combo_box(
                                    &request.combo_state,
                                    &format!(
                                        "Escolha um nome para a nova aba (e.g. \"70 - \
                                         {condo_name}\")",
                                    ),
                                    None,
                                    {
                                        let obs = obs.clone();
                                        move |assoc_sheet| {
                                            Message::ConfirmSheetAssoc(
                                                Some(obs.clone()),
                                                assoc_sheet,
                                                condo_name.clone(),
                                            )
                                        }
                                    }
                                )
                                .on_input(Message::SetAssocSheet),
                                button("Pular")
                                    .style(button_style())
                                    .height(32.0)
                                    .on_press(Message::SkipAssoc(obs))
                            ]
                            .spacing(10.0)
                            .align_x(Alignment::End)
                        }
                    })
                    .width(Length::Fill)
                }),

                None if self.began_processing => column.push(
                    main_container!(row![
                        text("Processando arquivos...")
                            .align_x(Alignment::Center)
                            .align_y(Alignment::Center)
                            .height(32.0)
                            .width(Length::Fill),
                        iced_aw::Spinner::new().height(32.0)
                    ])
                    .width(Length::Fill),
                ),
                None => column,
            },
        }
    }
}

#[derive(Clone)]
enum Message {
    FilePicked(&'static str, Option<FileHandle>),
    PickFile(&'static str, &'static str),
    StartRepassAddition,
    Result(Result<(), Arc<dyn std::error::Error + Send + Sync>>),
    SheetAssocRequested(Arc<Mutex<Option<ProcessParts>>>, AssocRequestReason),
    SkipAssoc(String),
    SetAssocSheet(String),
    ConfirmSheetAssoc(Option<String>, String, String),
    SetCondoName(String),
}

struct ProcessParts<I: Iterator<Item = Rec> = Box<dyn Iterator<Item = Rec> + Send + Sync>> {
    spreadsheet: Spreadsheet,
    financeiro: Vec<StringRecord>,
    vendas: I,
    repasses: PathBuf,
}

impl ProcessParts<StringRecordsIntoIter<std::fs::File>> {
    fn new(
        vendas: &Path,
        financeiro: &Path,
        repasses: &Path,
    ) -> Result<Self, Box<dyn std::error::Error + Send + Sync>> {
        fn csv_reader(
            path: impl AsRef<Path>,
        ) -> Result<csv::Reader<std::fs::File>, Box<dyn std::error::Error + Send + Sync>> {
            let mut reader = csv::ReaderBuilder::new();
            reader.delimiter(b';');
            Ok(reader.from_path(path)?)
        }

        Ok(Self {
            spreadsheet: umya_spreadsheet::reader::xlsx::lazy_read(repasses)?,
            financeiro: csv_reader(financeiro)?
                .into_records()
                .collect::<Result<Vec<_>, _>>()?,
            vendas: csv_reader(vendas)?.into_records(),
            repasses: repasses.to_path_buf(),
        })
    }
}

async fn process_csvs<I: Iterator<Item = Rec> + Send + Sync + 'static>(
    mut parts: ProcessParts<I>,
) -> Result<Option<(ProcessParts, AssocRequestReason)>, Box<dyn std::error::Error + Send + Sync>> {
    let sale_id = Regex::new(r"nº (\d{4,5})")?;

    let mut prev_id = None;
    let mut relevant_records = Vec::new();
    let mut honorarios = None;

    while let Some(record) = parts.vendas.next() {
        let record = record?;

        // Get the empty row right after all values are listed.
        if let Some(prev_id) = prev_id
            && record[NUMERO_DE_VENDA_VEN].is_empty()
        {
            let total_value_sum = parse_num(&record[VALOR_TOTAL_VEN])?;

            // Search for any record that has the same sale id.
            // If not found, then the sale is not registered.
            for paid_amount in parts.financeiro.iter().filter_map(|record| {
                sale_id
                    .captures(&record[OBSERVACOES_FIN])
                    .and_then(|captures| captures[1].parse::<usize>().ok())
                    .filter(|sale_id| *sale_id == prev_id)
                    .and_then(|_| parse_num(&record[VALOR_FIN]).ok())
            }) {
                match add_to_sheet(
                    &mut parts.spreadsheet,
                    paid_amount - honorarios.unwrap_or(0.0) * (paid_amount / total_value_sum),
                    total_value_sum - honorarios.unwrap_or(0.0),
                    &relevant_records,
                ) {
                    Ok(Some(reason)) => {
                        return Ok(Some((
                            ProcessParts {
                                spreadsheet: parts.spreadsheet,
                                financeiro: parts.financeiro,
                                vendas: Box::new(
                                    relevant_records
                                        .into_iter()
                                        .chain([record])
                                        .map(Ok)
                                        .chain(parts.vendas),
                                ),
                                repasses: parts.repasses,
                            },
                            reason,
                        )));
                    }
                    Ok(None) => {}
                    Err(err) => Err(err)?,
                }
            }

            honorarios = None;
            relevant_records.clear();
        } else if let Ok(sale_id) = record[NUMERO_DE_VENDA_VEN].parse::<usize>() {
            if unaccent(&record[DETALHAMENTO_VEN])
                .to_uppercase()
                .starts_with("HONORARIOS")
            {
                honorarios = Some(parse_num(&record[VALOR_TOTAL_VEN])?);
            } else {
                if let Some(prev_id) = prev_id
                    && prev_id != sale_id
                {
                    relevant_records.clear();
                }

                prev_id = Some(sale_id);
                relevant_records.push(record);
            }
        }
    }

    umya_spreadsheet::writer::xlsx::write(&parts.spreadsheet, &parts.repasses)?;

    Ok(None)
}

impl std::fmt::Display for AssocRequestReason {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            AssocRequestReason::NotFound(name) => f.write_str(name),
            AssocRequestReason::NewCondo(name) => f.write_str(name),
        }
    }
}

fn add_to_sheet(
    spreadsheet: &mut Spreadsheet,
    paid_amount: f64,
    total_value_sum: f64,
    records: &[StringRecord],
) -> Result<Option<AssocRequestReason>, Box<dyn std::error::Error + Send + Sync>> {
    fn round(value: f64) -> f64 {
        (value * 100.0).round() / 100.0
    }

    static ALREADY_ADDED: LazyLock<Mutex<HashMap<String, usize>>> = LazyLock::new(Mutex::default);
    let Some(first_rec) = records.first() else {
        return Ok(None);
    };

    let associations = ASSOCIATIONS.lock().unwrap();
    let name = {
        let name = unaccent(&first_rec[DESCRICAO_VEN]).to_uppercase();
        match name.strip_prefix("CONDOMINIO ") {
            Some(name) => name.to_string(),
            None => name,
        }
    };

    let name = match name.as_str() {
        "ROMA" => "ROMA BONFIM",
        "DELBOUX BLOCO B" => "DELBOUX B",
        "ROMA SAO JOAQUIM DA BARRA" => "ROMA S.J.BARRA",
        "MONTE ALEGRE" | "MONTE ALEGRE - BARRETOS" => "MONTE ALEGRE BARRETOS",
        name => name,
    };

    if name == "-" || name == "HONORARIO" {
        return Ok(None);
    }

    let (sheet, name) = if name == "NOVO CONDOMINIO" {
        if let Some(associated) = associations.get(&first_rec[OBSERVACAO_VEN]) {
            match associated {
                Some((assoc_sheet, condo_name)) => (
                    spreadsheet.get_sheet_by_name_mut(assoc_sheet).unwrap(),
                    condo_name.as_str(),
                ),
                None => return Ok(None),
            }
        } else {
            return Ok(Some(AssocRequestReason::NewCondo(
                first_rec[OBSERVACAO_VEN].to_string(),
            )));
        }
    } else {
        match spreadsheet
            .get_sheet_collection_no_check()
            .iter()
            .position(|sheet| {
                let name = if let Some(associated) = associations.get(name) {
                    match associated {
                        Some((assoc_name, _)) => assoc_name.as_ref(),
                        None => return false,
                    }
                } else {
                    name
                };

                unaccent(sheet.get_name())
                    .to_uppercase()
                    .ends_with(&unaccent(name).to_uppercase())
            })
            .and_then(|i| spreadsheet.get_sheet_mut(&i))
        {
            Some(sheet) => (sheet, name),
            None => {
                if let Some(None) = associations.get(name) {
                    return Ok(None);
                } else {
                    return Ok(Some(AssocRequestReason::NotFound(name.to_string())));
                }
            }
        }
    };

    // Add a new table in case this is the first time a certain sheet is
    // used.
    let mut already_added = ALREADY_ADDED.lock().unwrap();

    let base_style = {
        let mut border = umya_spreadsheet::Border::default();
        let mut color = Color::default();
        color.set_argb("000000");
        border.set_style(BorderStyleValues::Thin).set_color(color);

        let mut style = Style::get_default_value();
        style
            .get_borders_mut()
            .set_left(border.clone())
            .set_right(border.clone())
            .set_bottom(border.clone())
            .set_top(border);

        style
    };

    let values_count = {
        let key = unaccent(name).to_uppercase();

        if let Some(values_count) = already_added.get_mut(&key) {
            *values_count += records.len();
            *values_count
        } else {
            sheet.insert_new_row(&0, &6);

            let header_style = {
                let mut style = base_style.clone();
                let mut font = Font::default();
                font.set_size(20.0);

                style.set_background_color("9bbb59").set_font(font);
                style
            };

            let footer_style = {
                let mut style = header_style.clone();
                let mut font = Font::default();
                font.set_size(16.0);
                style.set_font(font);
                style
            };

            let complement_style = {
                let mut style = base_style.clone();
                style.set_background_color("dbe5f1");
                style
            };

            // Add header
            sheet.add_merge_cells("A1:B1");
            sheet.add_merge_cells("C1:E1");
            sheet.add_merge_cells("F1:H1");

            let date = chrono::Local::now().format("%m/%d/%Y").to_string();

            for (cell, value) in [("A1", "DATA DO REPASSE:"), ("C1", &date), ("F1", &key)] {
                sheet
                    .get_cell_mut(cell)
                    .set_value_string(value)
                    .set_style(header_style.clone());
            }

            let mut border = umya_spreadsheet::Border::default();
            border.set_style(BorderStyleValues::None);
            sheet
                .get_style_mut("C1")
                .get_borders_mut()
                .set_left(border.clone());

            // Add footer
            sheet.add_merge_cells("A5:G5");

            for (cell, value) in [("A5", "TOTAL DA TRANSFERÊNCIA"), ("H5", "")] {
                sheet
                    .get_cell_mut(cell)
                    .set_value_string(value)
                    .set_style(footer_style.clone());
            }

            sheet.get_cell_mut("H5").set_formula("H3-H4");

            // Add the other complementary cells
            sheet.add_merge_cells("A3:F3");
            sheet.add_merge_cells("A4:G4");

            for (cell, value) in [
                ("A2", "NUM."),
                ("B2", "UNIDADE"),
                ("C2", "COND."),
                ("D2", "PARCELA"),
                ("E2", "DESCRIÇÃO"),
                ("F2", "DETALHAMENTO"),
                ("G2", "RECEBIMENTOS"),
                ("H2", "TOTAL COBRANÇA"),
                ("A3", "TOTAL"),
                ("A4", "REEMBOLSO"),
            ] {
                sheet
                    .get_cell_mut(cell)
                    .set_value_string(value)
                    .set_style(complement_style.clone());
            }

            for cell in ["G3", "H3", "H4"] {
                sheet.get_cell_mut(cell).set_style(complement_style.clone());
            }

            *already_added
                .entry(key)
                .insert_entry(records.len())
                .into_mut()
        }
    };

    sheet.insert_new_row(&3, &(records.len() as u32));

    for (i, record) in records.iter().enumerate() {
        let details = &record[DETALHAMENTO_VEN];
        let total_value_portion = parse_num(&record[VALOR_TOTAL_VEN])?;
        let paid_portion = paid_amount * (total_value_portion / total_value_sum);

        sheet
            .get_cell_mut(format!("F{}", i + 3).as_str())
            .set_value_string(details);
        sheet
            .get_cell_mut(format!("G{}", i + 3).as_str())
            .set_value_number(round(paid_portion));
    }

    let record = &records[0];

    if records.len() > 1 {
        sheet.add_merge_cells(format!("E3:E{}", 2 + records.len()));
    }

    for (cell, index) in [
        ("A3", NUMERO_DE_VENDA_VEN),
        ("B3", UNIDADE_VEN),
        ("C3", DESCRICAO_VEN),
        ("D3", SITUACAO_VEN),
        ("E3", OBSERVACAO_VEN),
    ] {
        sheet.get_cell_mut(cell).set_value(&record[index]);
    }

    sheet
        .get_cell_mut(format!("H{}", 2 + records.len()))
        .set_value_number(round(paid_amount));

    sheet
        .get_cell_mut(format!("G{}", 3 + values_count).as_ref())
        .set_formula(format!("SUM(G3:G{})", 2 + values_count));
    sheet
        .get_cell_mut(format!("H{}", 3 + values_count).as_ref())
        .set_formula(format!("SUM(H3:H{})", 2 + values_count));

    sheet.set_style_by_range(format!("A3:A{}", 2 + records.len()).as_str(), {
        let mut style = base_style.clone();
        style.set_background_color("9bbb59");
        style
    });

    sheet.set_style_by_range(format!("B3:I{}", 2 + records.len()).as_str(), base_style);

    let alignment = sheet.get_style_mut("E3").get_alignment_mut();
    alignment.set_wrap_text(true);
    alignment.set_vertical(VerticalAlignmentValues::Top);

    Ok(None)
}

fn parse_num(num: impl AsRef<str>) -> Result<f64, ParseFloatError> {
    let num = num.as_ref();
    num.replace(".", "").replace(",", ".").parse()
}

fn button_style() -> impl Fn(&Theme, button::Status) -> button::Style {
    |theme, state| {
        let palette = Theme::extended_palette(theme);
        let mut style = button::Style::default().with_background(match state {
            button::Status::Active => palette.background.strong.color,
            button::Status::Hovered => palette.background.stronger.color,
            button::Status::Pressed => palette.background.strongest.color,
            button::Status::Disabled => palette.background.weaker.color,
        });
        style.border = Border {
            color: palette.background.strong.color,
            width: 0.0,
            radius: Radius::new(5.0),
        };
        style.text_color = match state {
            button::Status::Active | button::Status::Hovered | button::Status::Pressed => {
                style.text_color
            }
            button::Status::Disabled => palette.background.stronger.color,
        };
        style
    }
}

struct AssocRequest {
    parts: ProcessParts,
    reason: AssocRequestReason,
    combo_state: combo_box::State<String>,
    assoc_sheet: Option<String>,
    condo_name: String,
}

#[derive(Clone)]
enum AssocRequestReason {
    NotFound(String),
    NewCondo(String),
}

static ASSOCIATIONS: LazyLock<Mutex<Associations>> = LazyLock::new(Mutex::default);

const OBSERVACOES_FIN: usize = 5;
const VALOR_FIN: usize = 3;
const NUMERO_DE_VENDA_VEN: usize = 0;
const UNIDADE_VEN: usize = 1;
const DESCRICAO_VEN: usize = 2;
const SITUACAO_VEN: usize = 3;
const OBSERVACAO_VEN: usize = 4;
const DETALHAMENTO_VEN: usize = 5;
const VALOR_TOTAL_VEN: usize = 6;

type Rec = Result<StringRecord, csv::Error>;
type Associations = HashMap<String, Option<(String, String)>>;
