###############################
###  PIVOT TABLE FOR SHINY  ###
###############################
#' Title Rounds the count number
#'
#' @param cnt input number.
#' @param seed_n seed.
#'
#' @return round number
#' @export
#'
rRound <- function(cnt, seed_n) {
  set.seed(seed_n)
  prob <- runif(1, 0, 1)
  end_num<-cnt%%10
  len=length(cnt)
  for (i in 1:len) {
    if (end_num[i]==1|end_num[i]==6) {if (prob<=0.2) {cnt[i]=ceil(cnt[i]/5)*5} else {cnt[i]=floor(cnt[i]/5)*5}}
    else if (end_num[i]==2|end_num[i]==7) {if (prob<=0.4) {cnt[i]=ceil(cnt[i]/5)*5} else {cnt[i]=floor(cnt[i]/5)*5}}
    else if (end_num[i]==3|end_num[i]==8) {if (prob<=0.6) {cnt[i]=ceil(cnt[i]/5)*5} else {cnt[i]=floor(cnt[i]/5)*5}}
    else if (end_num[i]==4|end_num[i]==9) {if (prob<=0.8) {cnt[i]=ceil(cnt[i]/5)*5} else {cnt[i]=floor(cnt[i]/5)*5}}
    else {cnt[i]=cnt[i]}
  }
  return(as.integer(cnt))
}

get_expr2 <- function(idc, target, additional_expr,removal) {
  text_idc <- c(
    list(
      "Count" = "'n()'",
      "Count_distinct" = "paste0('n_distinct(', target, ', na.rm = TRUE)')",
      "Sum" = "paste0('sum(', target, ', na.rm = TRUE)')",
      "Mean" = "paste0('mean(', target, ', na.rm = TRUE)')",
      "Min" = "paste0('min(', target, ', na.rm = TRUE)')",
      "Max" = "paste0('max(', target, ', na.rm = TRUE)')",
      "Median" = "paste0('median(', target, ', na.rm = TRUE)')",
      "Variance" = "paste0('var(', target, ', na.rm = TRUE)')",
      "Standard_deviation" = "paste0('sd(', target, ', na.rm = TRUE)')"
    ),
    additional_expr)
  if(!is.null(removal))
  {
    for (i in 1:length(removal))
    {
      text_idc[names(text_idc)==removal[i]]=NULL
    }
  }
  return(eval(parse(text = text_idc[[idc]])))
}

toggleBtnSPivot <- function(session, inputId, type = "disable") {
  session$sendCustomMessage(
    type = "togglewidgetShinyPivot",
    message = list(inputId = inputId, type = type)
  )
}

combine_padding <- function(session, inputId, type = "regular") {
  session$sendCustomMessage(
    type = "method_combine_padding",
    message = list(inputId = inputId, type = type)
  )
}

extract_createTooltipOrPopoverOnUI <- function(id, type, options) {

  options = paste0("{'", paste(names(options), options, sep = "': '", collapse = "', '"), "'}")

  shiny::tags$script(shiny::HTML(paste0("$(document).ready(function() {setTimeout(function() {extract_shinyBS.addTooltip('", id, "', '", type, "', ", options, ")}, 500)});")))

}
extract_buildTooltipOrPopoverOptionsList <- function(title, placement, trigger, options, content) {

  if(is.null(options)) {
    options = list()
  }

  if(!missing(content)) {
    content <- gsub("'", "&#39;", content, fixed = TRUE)
    if(is.null(options$content)) {
      options$content = shiny::HTML(content)
    }
  }

  if(is.null(options$placement)) {
    options$placement = placement
  }

  if(is.null(options$trigger)) {
    if(length(trigger) > 1) trigger = paste(trigger, collapse = " ")
    options$trigger = trigger
  }

  if(is.null(options$title)) {
    options$title = title
    options$title <- gsub("'", "&#39;", options$title, fixed = TRUE)
  }

  return(options)

}



#' extract from shinyBS (extract_bsPopover)
#'
#' @param id Input id.
#' @param title The title of the popover.
#' @param content The main content of the popover.
#' @param placement Where the popover should appear relative to its target (top, bottom, left, or right). Defaults to bottom.
#' @param trigger What action should cause the popover to appear? (hover, focus, click, or manual). Defaults to hover.
#' @param options A named list of additional options to be set on the popover.
#'
#' @noRd
extract_bsPopover <- function (id, title, content, placement = "bottom", trigger = "hover",
                               options = NULL)
{
  options = extract_buildTooltipOrPopoverOptionsList(title, placement,
                                                     trigger, options, content)
  extract_createTooltipOrPopoverOnUI(id, "popover", options)
}

#' Shiny module to render and export pivot tables.
#'
#' @param input shiny input
#' @param output shiny input
#' @param session shiny input
#' @param id \code{character}. An ID string.
#' @param table_title \code{character}. An name used in the table.
#' @param headername \code{character}. An name used in Excel header when download the table.
#' @param footername \code{character}. An name used in Excel footer when download the table.
#' @param app_colors \code{character}. Vector of two colors c("#59bb28", "#217346") (borders)
#' @param remove_indicator \code{list} (NULL). The indicators do no use.
#' @param app_linewidth \code{numeric}. Borders width
#' @param data \code{data.frame} / \code{data.table}. Initial data table.s
#' @param pivot_cols \code{character} (NULL). Columns to be used as pivot in rows and cols.
#' @param max_n_pivot_cols \code{numeric} (100). Maximum unique values for a \code{pivot_cols} if pivot_cols = NULL
#' @param indicator_cols \code{character} (NULL). Columns on which indicators will be calculated.
#' @param additional_expr_num \code{named list} (list()). Additional computations to be allowed for quantitative vars.
#' @param additional_expr_char \code{named list} (list()). Additional computations to be allowed for qualitative vars.
#' @param additional_combine \code{named list} (list()). Additional combinations to be allowed.
#' @param theme \code{list} (NULL). Theme to customize the output of the pivot table.
#' @param export_styles \code{boolean} (TRUE). Whether or not to apply styles (like the theme) when exporting to Excel.
#' @param show_title \code{boolean} (TRUE). Whether or not to display the app title.
#' Some styles may not be supported by Excel.
#' @param initialization  \code{named list} (NULL). Initialization parameters to display a table when launching the module.
#' Available fields are :
#'\itemize{
#'  \item{\code{rows:}}{ Selected pivot rows.}
#'  \item{\code{cols:}}{ Selected pivot columns.}
#'  \item{\code{target, combine target:}} { Selected target and combine_target columns.}.
#'  \item{\code{idc, combine_idc:}}{ Selected idc and combine_idc columns.}
#'  \item{\code{combine:}}{ Selected combine operator.}
#'  \item{\code{format_digit, format_prefix, format_suffix, format_sep_thousands, format_decimal:}}{ Selected formats for the table idc.}
#'  \item{\code{idcs:}}{ idcs to be displayed (list of named list), see the example to get the fields.}
#'}
#'
#' @return Nothing. Just Start a Shiny module.
#'
#' @import pivottabler shiny openxlsx
#' @importFrom colourpicker colourInput
#'
#' @export
#' @rdname shiny_pivot_table


shinypivottabler2 <- function(input, output, session,
                              data,
                              headername = "",
                              footername = "",
                              pivot_cols = NULL,
                              indicator_cols = NULL,
                              remove_indicator=NULL,
                              max_n_pivot_cols = 100,
                              additional_expr_num = list(),
                              additional_expr_char = list(),
                              additional_combine = list(),
                              theme = NULL,
                              export_styles = TRUE,
                              show_title = TRUE,
                              initialization = NULL) {

  ns <- session$ns

  observe({
    if (! is.null(idcs()) && length(idcs()) > 0) {
      toggleBtnSPivot(session = session, inputId = ns("go_table"), type = "enable")
    } else {
      toggleBtnSPivot(session = session, inputId = ns("go_table"), type = "disable")
    }
  })
  observe({
    if (! is.null(input$combine) && input$combine != "None") {
      combine_padding(session = session, inputId = ns("id_padding_1"), type = "combine")
      combine_padding(session = session, inputId = ns("id_padding_2"), type = "combine")
      combine_padding(session = session, inputId = ns("id_padding_3"), type = "combine")
      combine_padding(session = session, inputId = ns("id_padding_4"), type = "combine")
    } else {
      combine_padding(session = session, inputId = ns("id_padding_1"), type = "regular")
      combine_padding(session = session, inputId = ns("id_padding_2"), type = "regular")
      combine_padding(session = session, inputId = ns("id_padding_3"), type = "regular")
      combine_padding(session = session, inputId = ns("id_padding_4"), type = "regular")
    }
  })

  # reactive controls
  if (! shiny::is.reactive(data)) {
    get_data <- shiny::reactive(data)
  } else {
    get_data <- data
  }

  have_data <- reactive({
    data <- get_data()
    ! is.null(data) && any(c("data.frame", "tbl", "tbl_df", "data.table") %in% class(data)) && nrow(data) > 0
  })
  output$ui_have_data <- reactive({
    have_data()
  })
  outputOptions(output, "ui_have_data", suspendWhenHidden = FALSE)

  if (! shiny::is.reactive(pivot_cols)) {
    get_pivot_cols <- shiny::reactive(pivot_cols)
  } else {
    get_pivot_cols <- pivot_cols
  }

  if (! shiny::is.reactive(indicator_cols)) {
    get_indicator_cols <- shiny::reactive(indicator_cols)
  } else {
    get_indicator_cols <- indicator_cols
  }

  if (! shiny::is.reactive(max_n_pivot_cols)) {
    get_max_n_pivot_cols <- shiny::reactive(max_n_pivot_cols)
  } else {
    get_max_n_pivot_cols <- max_n_pivot_cols
  }

  get_theme <- reactiveVal(NULL)
  observe({
    if (! shiny::is.reactive(theme)) {
      if (is.null(theme)) {
        get_theme(list(
          fontName="Courier New, Courier",
          fontSize="1.2em",
          headerBackgroundColor = "#217346",
          headerColor = "rgb(255, 255, 255)",
          cellBackgroundColor = "rgb(255, 255, 255)",
          cellColor = "rgb(0, 0, 0)",
          outlineCellBackgroundColor = "rgb(192, 192, 192)",
          outlineCellColor = "rgb(0, 0, 0)",
          totalBackgroundColor = "#59bb28",
          totalColor = "rgb(0, 0, 0)",
          borderColor = "rgb(64, 64, 64)"))
      } else {
        get_theme(theme)
      }
    } else {
      if (is.null(theme())) {
        get_theme(list(
          fontName="Courier New, Courier",
          fontSize="1.2em",
          headerBackgroundColor = "#217346",
          headerColor = "rgb(255, 255, 255)",
          cellBackgroundColor = "rgb(255, 255, 255)",
          cellColor = "rgb(0, 0, 0)",
          outlineCellBackgroundColor = "rgb(192, 192, 192)",
          outlineCellColor = "rgb(0, 0, 0)",
          totalBackgroundColor = "#59bb28",
          totalColor = "rgb(0, 0, 0)",
          borderColor = "rgb(64, 64, 64)"))
      } else {
        get_theme(theme())
      }
    }
  })

  if (! shiny::is.reactive(export_styles)) {
    get_export_styles <- shiny::reactive(export_styles)
  } else {
    get_export_styles <- export_styles
  }

  if (! shiny::is.reactive(show_title)) {
    get_show_title <- shiny::reactive(show_title)
  } else {
    get_show_title <- show_title
  }

  if (! shiny::is.reactive(additional_expr_num)) {
    get_additional_expr_num <- shiny::reactive(additional_expr_num)
  } else {
    get_additional_expr_num <- additional_expr_num
  }

  if (! shiny::is.reactive(additional_expr_char)) {
    get_additional_expr_char <- shiny::reactive(additional_expr_char)
  } else {
    get_additional_expr_char <- additional_expr_char
  }

  if (! shiny::is.reactive(additional_combine)) {
    get_additional_combine <- shiny::reactive(additional_combine)
  } else {
    get_additional_combine <- additional_combine
  }

  if (! shiny::is.reactive(initialization)) {
    trigger_initialization <- shiny::reactive(initialization)
  } else {
    trigger_initialization <- shiny::reactive(initialization())
  }

  get_initialization <- reactiveVal(NULL)
  observe({
    if (! shiny::is.reactive(initialization)) {
      get_initialization(initialization)
    } else {
      get_initialization(initialization())
    }
  })

  output$show_title <- reactive({
    get_show_title()
  })
  outputOptions(output, "show_title", suspendWhenHidden = FALSE)

  # estimate the number of rows and cols in the pivot table
  ctrl_var_len <- reactive({
    sapply(get_data(), function(x) length(unique(x)))
  })

  output$estimated_size <- renderText({
    rows <- input$rows
    cols <- input$cols

    isolate({
      paste0("<b>Estimated size : ", ifelse(is.null(rows), 1, Reduce("*", ctrl_var_len()[rows])),
             "</b> rows  x  <b>",
             ifelse(is.null(cols), 1, Reduce("*", ctrl_var_len()[cols])), "</b> colums +  <b> Subtotals </b>")
    })
  })

  # update inputs
  observe({
    data <- get_data()
    trigger_initialization()

    isolate({
      pivot_cols <- get_pivot_cols()
      initialization <- get_initialization()

      if (is.null(pivot_cols)) {
        choices <- names(ctrl_var_len())[ctrl_var_len() <= get_max_n_pivot_cols()]

        updateSelectInput(session = session, "rows",
                          choices = c("", choices),
                          selected = if (is.null(initialization$rows)) {""} else {initialization$rows})
        updateSelectInput(session = session, "cols",
                          choices = c("", choices),
                          selected = if (is.null(initialization$cols)) {""} else {initialization$cols})
      } else {
        updateSelectInput(session = session, "rows",
                          choices = c("", pivot_cols),
                          selected = if (is.null(initialization$rows)) {""} else {initialization$rows})
        updateSelectInput(session = session, "cols",
                          choices = c("", pivot_cols),
                          selected = if (is.null(initialization$cols)) {""} else {initialization$cols})
      }
    })
  })

  observe({
    data <- get_data()
    trigger_initialization()

    isolate({
      indicator_cols <- get_indicator_cols()
      initialization <- get_initialization()

      if (is.null(indicator_cols) && have_data()) {
        updateSelectInput(session = session, "target",
                          choices = c("", names(which(sapply(data, function(x) any(class(x) %in% c("logical", "numeric", "integer", "character", "factor")))))),
                          #selected = if (is.null(initialization$target)) {""} else {initialization$target})
                          selected = if (is.null(initialization$target)) {"Count"} else {initialization$target})

      } else {
        updateSelectInput(session = session, "target",
                          choices = c("", indicator_cols),
                          #selected = if (is.null(initialization$target)) {""} else {initialization$target})
                          selected = if (is.null(initialization$target)) {"Count"} else {initialization$target})
      }
    })
  })

  observe({
    target <- input$target
    trigger_initialization()

    isolate({

      req(target)

      initialization <- get_initialization()

      if (is.null(get_data()[[target]]) || is.numeric(get_data()[[target]])) {
        choices <- sort(c(
          c("Count", "Count distinct", "Sum", "Mean", "Min", "Max", "Median", "Variance", "Standard deviation"),
          names(get_additional_expr_num())
        ))
        if (!is.null(remove_indicator)) {
          for (i in 1:length(remove_indicator)) {
            choices=choices[choices!=remove_indicator[i]]}
        }
        updateSelectInput(session = session, "idc",
                          choices = choices,
                          selected = if (is.null(initialization$idc)) {ifelse(input$idc %in% choices, input$idc, "Count")} else {initialization$idc})
      } else if (is.character(get_data()[[target]]) || is.factor(get_data()[[target]])) {
        choices <- sort(c(
          c("Count", "Count distinct"),
          names(get_additional_expr_char())
        ))
        for (i in 1:length(remove_indicator))
        {if (remove_indicator=="Count"||"Count distinct")
        {choices=choices[choices!=remove_indicator[i]]}
        }
        updateSelectInput(session = session, "idc",
                          choices = choices,
                          selected = if (is.null(initialization$idc)) {ifelse(input$idc %in% choices, input$idc, "Count")} else {initialization$idc})
      }
    })
  })

  observe({
    trigger_initialization()

    isolate({
      initialization <- get_initialization()

      updateSelectInput(session = session, "combine",
                        #                        choices = c(c("None" = "None",
                        #                                      "Add" = "+",
                        #                                      "Substract" = "-",
                        #                                      "Multiply" = "*",
                        #                                      "Divise" = "/"),
                        #                                    get_additional_combine()),
                        choices = ("None" = "None"),
                        selected = if (is.null(initialization$combine)) {"None"} else {initialization$combine})
    })
  })

  store_format <- reactiveValues("format_digit" = 0,
                                 "format_prefix" = "",
                                 "format_suffix" = "",
                                 "format_sep_thousands" = ",",
                                 "format_decimal" = ".")

  observe({
    cpt <- input$specify_format

    isolate({
      if (! is.null(cpt) && cpt > 0) {
        showModal(
          modalDialog(
            title = "Format the cells",
            fluidRow(
              column(4,
                     numericInput(ns("format_digit"), label = "Nb. digits",
                                  min = 0, max = Inf, value = store_format[["format_digit"]], step = 1, width = "100%")
              ),
              column(4,
                     textInput(ns("format_prefix"), label = "Prefix (excel only)",
                               value = store_format[["format_prefix"]], width = "100%")
              ),
              column(4,
                     textInput(ns("format_suffix"), label = "Suffix (excel only)",
                               value = store_format[["format_suffix"]], width = "100%")
              )
            ),
            fluidRow(
              column(4,
                     selectInput(ns("format_sep_thousands"), label = "Thousands sep.",
                                 choices = c("None", "Space" = " ", ","), selected = store_format[["format_sep_thousands"]], width = "100%")
              ),
              column(4,
                     selectInput(ns("format_sep_decimals"), label = "Decimal sep.",
                                 choices = c(".", ","), selected = store_format[["format_decimal"]], width = "100%")
              )
            ),
            easyClose = FALSE,
            footer = div(style = "margin-right: 20px;",
                         fluidRow(
                           column(3,
                                  div(actionButton(inputId = ns("format_valid"), label = "Validate", width = "100%"), align = "left")
                           ),
                           column(3, offset = 6,
                                  div(actionButton(inputId = ns("format_cancel"), label = "Cancel", width = "100%"), align = "right")
                           ))
            ))
        )
      }
    })
  })
  observe({
    cpt_valid <- input$format_valid
    cpt_cancel <- input$format_cancel

    isolate({
      if (! is.null(cpt_valid) && cpt_valid > 0) {
        store_format$format_digit <- input$format_digit
        store_format$format_prefix <- input$format_prefix
        store_format$format_suffix <- input$format_suffix
        store_format$format_sep_thousands <- input$format_sep_thousands
        store_format$format_decimal <- input$format_sep_decimals

        shiny::removeModal()
      }
      if (! is.null(cpt_cancel) && cpt_cancel > 0) {
        updateNumericInput(session = session, "format_digit",
                           value = store_format[["format_digit"]])
        updateTextInput(session = session, "format_prefix",
                        value = store_format[["format_digit"]])
        updateTextInput(session = session, "format_suffix",
                        value = store_format[["format_suffix"]])
        updateSelectInput(session = session, "format_sep_thousands",
                          selected = store_format[["format_sep_thousands"]])
        updateSelectInput(session = session, "format_sep_decimals",
                          selected = store_format[["format_decimal"]])

        shiny::removeModal()
      }
    })
  })

  idcs <- reactiveVal()


  observe({
    trigger_initialization()

    isolate({
      initialization <- get_initialization()

      if (! is.null(initialization)) {
        # update format defaults
        store_format$format_digit <- ifelse(is.null(initialization$format_digit), store_format$format_digit, initialization$format_digit)
        store_format$format_prefix <- ifelse(is.null(initialization$format_prefix), store_format$format_prefix, initialization$format_prefix)
        store_format$format_suffix <- ifelse(is.null(initialization$format_suffix), store_format$format_suffix, initialization$format_suffix)
        store_format$format_sep_thousands <- ifelse(is.null(initialization$format_sep_thousands), store_format$format_sep_thousands, initialization$format_sep_thousands)
        store_format$format_decimal <- ifelse(is.null(initialization$format_decimal), store_format$format_decimal, initialization$format_decimal)

        # update idcs
        if (! is.null(initialization$idcs)) {
          idcs(initialization$idcs)
        }
      }
    })
  })


  store_pt <- reactiveVal(NULL)
  observe({
    cpt <- input$go_table
    trigger_initialization()

    isolate({
      idcs <- isolate(idcs())
      data <- isolate(get_data())
      initialization <- get_initialization()

      if (! is.null(data) && (((! is.null(cpt) && cpt > 0 && ! is.null(idcs)) || ! is.null(initialization)) && length(idcs) > 0)) {
        shiny::withProgress(message = 'Creating the table...', value = 0.5, {

          names(data) <- gsub("[[:punct:]| ]", "_", names(data))
          pt <- pivottabler::PivotTable$new()
          pt$addData(data)

          # rows and columns
          rows <- if (is.null(initialization$rows)) {input$rows} else {initialization$rows}
          for (row in rows) {
            if (! is.null(row) && row != "") {pt$addRowDataGroups(gsub("[[:punct:]| ]", "_", row))}
          }
          cols <- if (is.null(initialization$cols)) {input$cols} else {initialization$cols}
          for (col in cols) {
            if (! is.null(col) && col != "") {pt$addColumnDataGroups(gsub("[[:punct:]| ]", "_", col))}
          }

          for (index in 1:length(idcs)) {
            is_checked <- input[[paste0("idc_name_box_", index)]]


            label <- idcs[[index]][["label"]]
            #target <- gsub("[[:punct:]| ]", "_", idcs[[index]]["target"])
            target <- gsub(" ", "_", idcs[[index]]["target"])
            idc <- gsub(" ", "_", idcs[[index]][["idc"]])
            nb_decimals <- ifelse(is.na(idcs[[index]]["nb_decimals"]), 1, idcs[[index]]["nb_decimals"])
            sep_thousands <- ifelse(is.na(idcs[[index]]["sep_thousands"]), " ", idcs[[index]]["sep_thousands"])
            sep_decimal <- ifelse(is.na(idcs[[index]]["sep_decimal"]), ".", idcs[[index]]["sep_decimal"])
            prefix <- ifelse(is.na(idcs[[index]]["prefix"]), "", idcs[[index]]["prefix"])
            suffix <- ifelse(is.na(idcs[[index]]["suffix"]), "", idcs[[index]]["suffix"])
            combine <- if ("combine" %in% names(idcs[[index]])) {idcs[[index]]["combine"]} else {NULL}
            combine_target <- if ("combine_target" %in% names(idcs[[index]])) {gsub("[[:punct:]| ]", "_", idcs[[index]]["combine_target"])} else {NULL}
            combine_idc <- if ("combine_idc" %in% names(idcs[[index]])) {gsub(" ", "_", idcs[[index]][["combine_idc"]])} else {NULL}
            pt$defineCalculation(calculationName = paste0(target, "_", tolower(idc), "_", index),
                                 caption = label,
                                 summariseExpression = get_expr2(idc, target, additional_expr = c(get_additional_expr_num(), get_additional_expr_char()), removal=remove_indicator),
                                 format = list("digits" = nb_decimals, "nsmall" = nb_decimals,
                                               "decimal.mark" = sep_decimal,
                                               "big.mark" = ifelse(sep_thousands == "None", "", sep_thousands),
                                               scientific = T),
                                 cellStyleDeclarations = list("xl-value-format" = paste0(prefix, ifelse(sep_thousands == "None", "", paste0("#", sep_thousands)), "##0", ifelse(nb_decimals > 0, paste0(sep_decimal, paste0(rep(0, nb_decimals), collapse = "")), ""), suffix)),
                                 visible = ifelse(is.null(combine_target), T, F))
            if (! is.null(combine_target) && combine_target != "") {
              pt$defineCalculation(calculationName = paste0(combine_target, "_", tolower(combine_idc), "_combine_", index),
                                   summariseExpression = get_expr2(combine_idc, combine_target, additional_expr = c(get_additional_expr_num(), get_additional_expr_char()), removal=remove_indicator),
                                   visible = FALSE)
              pt$defineCalculation(calculationName = paste0(combine_target, "_", tolower(combine_idc), combine, combine_target, "_", tolower(combine_idc), "_combine_", index),
                                   caption = label,
                                   basedOn = c(paste0(target, "_", tolower(idc), "_", index), paste0(combine_target, "_", tolower(combine_idc), "_combine_", index)),
                                   type = "calculation",
                                   calculationExpression = paste0("values$", paste0(target, "_", tolower(idc), "_", index), combine, "values$", paste0(combine_target, "_", tolower(combine_idc), "_combine_", index)),
                                   format = list("digits" = nb_decimals, "nsmall" = nb_decimals,
                                                 "decimal.mark" = sep_decimal,
                                                 "big.mark" = ifelse(sep_thousands == "None", "", sep_thousands), scientific = F),
                                   cellStyleDeclarations = list("xl-value-format" = paste0(prefix, ifelse(sep_thousands == "None", "", paste0("#", sep_thousands)), "##0", ifelse(nb_decimals > 0, paste0(sep_decimal, paste0(rep(0, nb_decimals), collapse = "")), ""), suffix)))
            }
          }

          ctrl <- tryCatch(pt$evaluatePivot(), error = function(e){
            showModal(modalDialog(
              title = "Error creating pivot table",
              e$message,
              easyClose = TRUE,
              footer = NULL
            ))
            "error"
          })

          if(!isTRUE(all.equal(ctrl, "error"))){
            store_pt(pt)
          } else {
            store_pt(NULL)
          }

        })
      } else {
        store_pt(NULL)
      }
    })
  })

  observe({
    cpt <- input$update_theme

    isolate({
      theme <- get_theme()

      if (! is.null(cpt) && cpt > 0) {
        showModal(
          modalDialog(
            title = "Update the theme",

            fluidRow(
              column(12,
                     textInput(ns("theme_fontname"), label = "Font name",
                               value = theme$fontName),
                     numericInput(ns("theme_fontsize"), label = "Font size (em)",
                                  value = as.numeric(gsub("em$", "", theme$fontSize)), min = 0, max = 10, step = 0.5),
                     column(6,
                            colourpicker::colourInput(ns("theme_headerbgcolor"), label = "Header bg color",
                                                      value = theme$headerBackgroundColor)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_headercolor"), label = "Header text color",
                                                      value = theme$headerColor)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_cellbgcolor"), label = "Cell bg color",
                                                      value = theme$cellBackgroundColor)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_cellcolor"), label = "Cell text color",
                                                      value = theme$cellColor)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_outlinecellbgcolor"), label = "Outline cell bg color",
                                                      value = theme$outlineCellBackgroundColor)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_outlinecellcolor"), label = "Outline text cell color",
                                                      value = theme$OutlineCell)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_totalbgcolor"), label = "Total bg color",
                                                      value = theme$totalBackgroundColor)
                     ),
                     column(6,
                            colourpicker::colourInput(ns("theme_totalcolor"), label = "Total text color",
                                                      value = theme$totalColor)
                     ),
                     colourpicker::colourInput(ns("theme_bordercolor"), label = "Border color",
                                               value = theme$borderColor)
              )
            ),
            easyClose = FALSE,
            footer = div(style = "margin-right: 20px;",
                         fluidRow(
                           column(3,
                                  div(actionButton(inputId = ns("theme_valid"), label = "Validate", width = "100%"), align = "left")
                           ),
                           column(3, offset = 6,
                                  div(actionButton(inputId = ns("theme_cancel"), label = "Cancel", width = "100%"), align = "right")
                           ))
            ))
        )
      }
    })
  })

  observe({
    cpt_valid <- input$theme_valid
    cpt_cancel <- input$theme_cancel

    isolate({

      if (! is.null(cpt_valid) && ! is.null(cpt_cancel) && (cpt_valid > 0 || cpt_cancel > 0)) {
        if (cpt_valid > 0) {
          theme <- get_theme()

          theme$fontName <- input$theme_fontname
          theme$fontSize <- paste0(input$theme_fontsize, "em")
          theme$headerBackgroundColor <- input$theme_headerbgcolor
          theme$headerColor <- input$theme_headercolor
          theme$cellBackgroundColor <- input$theme_cellbgcolor
          theme$cellColor <- input$theme_cellcolor
          theme$outlineCellBackgroundColor <- input$theme_outlinecellbgcolor
          theme$outlineCellColor <- input$theme_outlinecellcolor
          theme$totalBackgroundColor <- input$theme_totalbgcolor
          theme$totalColor <- input$theme_totalcolor
          theme$borderColor <- input$theme_bordercolor

          get_theme(theme)
        }

        shiny::removeModal()
      }
    })
  })

  counter_pivottable <- reactiveVal(0)

  observe({
    input$go_table
    theme <- get_theme()
    trigger_initialization()

    isolate({
      counter_pivottable(counter_pivottable() + 1)
      output[[paste0("pivottable_", counter_pivottable())]] <- renderPivottabler({

        isolate({
          pt <- store_pt()

          if (! is.null(pt)) {
            if (! is.null(get_initialization())) {
              get_initialization(NULL)
            }

            pt$theme <- theme

            pt$renderPivot()

          } else {
            NULL
          }
        })
      })
    })
  })

  output$pivottable <- renderUI({
    div(pivottablerOutput(ns(paste0("pivottable_", counter_pivottable())), width = "100%", height = "100%"), style = "padding-top: 1.5%;")
  })

  output$is_pivottable <- reactive({
    ! is.null(store_pt())
  })
  outputOptions(output, "is_pivottable", suspendWhenHidden = FALSE)

  get_wb <- reactive({
    pt <- store_pt()

    isolate({
      if (! is.null(pt)) {
        shiny::withProgress(message = 'Preparing the export...', value = 0.5, {
          wb <- createWorkbook(creator = "Shiny pivot table")
          addWorksheet(wb=wb, sheetName="Pivot table", header=c(headername, "", ""),footer=c(footername, "", ""))
          writeData(wb=wb, sheet="Pivot table", x= headername, startRow = 1)
          writeData(wb=wb, sheet="Pivot table", x= footername, startRow = 2)

          pt$writeToExcelWorksheet(wb = wb, wsName = "Pivot table",
                                   topRowNumber = 4, leftMostColumnNumber = 1,
                                   outputValuesAs = "formattedValueAsNumber",
                                   applyStyles = get_export_styles(), mapStylesFromCSS = TRUE)
          wb
        })
      }
    })
  })

  output$export <- downloadHandler(
    filename = function() {
      paste0("pivot_table_", base::format(Sys.time(), format = "%Y%m%d_%H%M%S") ,".xlsx")
    },
    content = function(file) {
      saveWorkbook(get_wb(), file = file, overwrite = TRUE)
    }
  )
}



#' @import pivottabler shiny
#'
#' @export
#'
#' @rdname shiny_pivot_table
#'
shinypivottablerUI2 <- function(id,
                                table_title = "Build your own pivot table",
                                app_colors = c("#59bb28", "#217346"),
                                app_linewidth = 8) {
  ns <- shiny::NS(id)

  fluidPage(
    conditionalPanel(condition = paste0("output['", ns("show_title"), "']"),
                     div(h2(HTML(paste0("<b> ", table_title, "<b>"))), style = paste0("color: ", app_colors[2], "; margin-left: 15px;"))
    ),

    br(),

    # tags
    singleton(tags$head(
      tags$script(src = "shiny_pivot_table/shinypivottable.js")
    )),
    singleton(tags$head(
      tags$link(rel = "stylesheet", type = "text/css", href = "shiny_pivot_table/shinypivottable.css")
    )),
    tags$head(
      tags$style(HTML("
        div.combine_padding { padding-top: 35px; }
        ")
      )
    ),

    conditionalPanel(condition = paste0("output['", ns("ui_have_data"), "']"),
                     fluidRow(style = "padding-left: 1%; padding-right: 1%;",
                              column(12, style = paste0("border-radius: 3px; border-top: ", app_linewidth, "px solid ", app_colors[1], "; border-bottom: ", app_linewidth, "px solid ", app_colors[2], "; border-left: ", app_linewidth, "px solid ", app_colors[1], "; border-right: ", app_linewidth, "px solid ", app_colors[2], ";"),

                                     br(),

                                     fluidRow(style = "margin-left: 0px; margin-bottom: -15px;",
                                              fluidRow(
                                                column(4,
                                                       selectInput(ns("rows"), label = "Please select your table rows",
                                                                   choices = NULL, multiple = T, width = "100%")
                                                ),
                                                column(4,
                                                       selectInput(ns("cols"), label = "Please select your table columns",
                                                                   choices = NULL, multiple = T, width = "100%")
                                                ),
                                                column(4,
                                                       div(htmlOutput(ns("estimated_size")), style = "margin-left: 20px; margin-top: 25px;")
                                                )
                                              )
                                     ),

                                     br()
                              )
                     ),

                     br(),
                     br(),

                     fluidRow(style = "padding-left: 1%; padding-right: 1%;",
                              column(12, style = paste0("padding: 2.5%; overflow-x: auto; overflow-y: auto; border-radius: 3px; border-top: ", app_linewidth, "px solid ", app_colors[1], "; border-bottom: ", app_linewidth, "px solid ", app_colors[2], "; border-left: ", app_linewidth, "px solid ", app_colors[1], "; border-right: ", app_linewidth, "px solid ", app_colors[2], ";"),
                                     fluidRow(
                                       #column(4, offset = 1,
                                       column(4, offset = 2,
                                              div(actionButton(ns("go_table"), label = "Display table", width = "100%" ), align = "right")
                                       ),
                                       #column(2,
                                       column(4,
                                              div(actionButton(ns("update_theme"), label = "Update theme", width = "100%" ), align = "right")
                                       )#,
                                       #column(4,
                                       #       div(actionButton(ns("reset_table"), label = "Reset table", width = "100%"), align = "left")
                                       # )
                                     ),

                                     br(),

                                     conditionalPanel(condition = paste0("output['", ns("is_pivottable"), "']"),
                                                      uiOutput(ns("pivottable")),
                                                      br(),
                                                      column(6, offset = 3,
                                                             div(downloadButton(ns("export"), label = "Download table"), align = "center", style = "width: 100%;")
                                                      )
                                     ),
                                     conditionalPanel(condition = paste0("! output['", ns("is_pivottable"), "']"),
                                                      div(h3("No data to display"), align = "center", style = paste0("color: ", app_colors[2], ";"))
                                     )
                              )
                     )
    ),
    conditionalPanel(condition = paste0("output['", ns("ui_have_data"), "'] === false"),
                     div(h3("No data to display"), align = "center", style = paste0("color: ", app_colors[2], ";"))
    )
  )
}
