# -*- coding: utf-8 -*-
# Axlsx is a gem or generating excel spreadsheets with charts, images and many other features.
#
# acts_as_xlsx provides integration into active_record for Axlsx.
#
require 'axlsx'

# Adding to the Axlsx module
# @see http://github.com/randym/axlsx
module Axlsx
  # === Overview
  # This module defines the acts_as_xlsx class method and provides to_xlsx support to both AR classes and instances
  module Ar
    def self.included(base) # :nodoc:
      base.send :extend, ClassMethods
    end

    # Class methods for the mixin
    module ClassMethods
      # defines the class method to inject to_xlsx
      # @option options [Array, Symbol] columns an array of symbols defining the columns and methods to call in generating sheet data for each row.
      # @option options [String] i18n (default nil) The path to search for localization. When this is specified your i18n.t will be used to determine the labels for columns.
      # @example
      #       class MyModel < ActiveRecord::Base
      #          acts_as_xlsx :columns=> [:id, :created_at, :updated_at], :i18n => 'activerecord.attributes'
      def acts_as_xlsx(options = {})
        cattr_accessor :xlsx_i18n, :xlsx_columns
        self.xlsx_i18n = options.delete(:i18n) || false
        begin
          self.xlsx_columns = options.delete(:columns) || column_names.map { |c| c = c.to_sym }
        rescue ActiveRecord::StatementInvalid
        end
        extend Axlsx::Ar::SingletonMethods
      end
    end

    # Singleton methods for the mixin
    module SingletonMethods
      # Maps the AR class to an Axlsx package
      # options are passed into AR find
      # @param [Array, Array] columns as an array of symbols or a symbol that defines the attributes or methods to render in the sheet.
      # @option options [Integer] header_style to apply to the first row of field names
      # @option options [Array, Symbol] types an array of Axlsx types for each cell in data rows or a single type that will be applied to all types.
      # @option options [Integer, Array] style The style to pass to Worksheet#add_row
      # @option options [String] i18n The path to i18n attributes. (usually activerecord.attributes)
      # @option options [Package] package An Axlsx::Package. When this is provided the output will be added to the package as a new sheet.
      # @option options [String] name This will be used to name the worksheet added to the package. If it is not provided the name of the table name will be humanized when i18n is not specified or the I18n.t for the table name.
      # @see Worksheet#add_row
      def to_xlsx(options = {})
        row_style = options.delete(:style)
        header_style = options.delete(:header_style) || row_style
        types = [options.delete(:types) || []].flatten

        i18n = options.delete(:i18n) || xlsx_i18n
        columns = options.delete(:columns) || xlsx_columns
        labels = options.delete(:labels) || {}
        as_array = options.delete(:as_array)

        p = options.delete(:package) || Package.new
        row_style = p.workbook.styles.add_style(row_style) unless row_style.nil?
        header_style = p.workbook.styles.add_style(header_style) unless header_style.nil?
        i18n = xlsx_i18n == true ? 'activerecord.attributes' : i18n
        sheet_name = options.delete(:name) || (i18n ? I18n.t("#{i18n}.#{table_name.underscore}") : table_name.humanize)
        data = options.delete(:data) || where(options)
        data = data.to_a if as_array

        return p if (data.try(:unscope, :group) || data).empty?

        bool_columns = []
        types.each_with_index.select { |t, i| t == :boolean }.map(&:last).each do |idx|
          types[idx] = :string
          bool_columns << columns[idx]
        end
        false_values = [0, false, nil, '0', 'false', '']

        if !row_style
          time_xf = p.workbook.styles.add_style(format_code: 'dd-mm-yyyy hh:mm:ss')
          nowrap_xf = p.workbook.styles.add_style(alignment: { wrap_text: true })
          date_xf = p.workbook.styles.add_style(format_code: 'dd-mm-yyyy')
          row_style = types.map do |type|
            case type
            when :time then time_xf
            when :string, nil then nowrap_xf
            when :date then date_xf
            end
          end
        end

        p.workbook.add_worksheet(name: sheet_name) do |sheet|
          col_labels = columns.map do |c|
            default = c.to_s.tr('.', '_').humanize
            if labels[c]
              labels[c]
            elsif i18n
              I18n.t("#{i18n}.#{name.underscore}.#{c}", default: default)
            else
              default
            end
          end

          sheet.add_row col_labels, style: header_style

          iterator = data.respond_to?(:find_each) ? [:find_each, { batch_size: 500 }] : [:each]
          data.send(*iterator) do |r|
            row_data = columns.each_with_index.map do |c, idx|
              value = r
              c.to_s.split(/(?<!\\)\./).each do |method|
                method = method.gsub('\.', '.')
                value = value.try(method) || value.try(:[], method)
              end
              if c.in?(bool_columns)
                value = value.in?(false_values) ? 'no' : 'yes'
                value = I18n.t("#{i18n}.generic.#{value}", default: value.titleize)
              end
              value = value.to_date if types[idx] == :date && value && !value.is_a?(Date)
              value
            end
            sheet.add_row row_data, style: row_style, types: types
          end
        end
        p
      end
    end
  end
end

require 'active_record'
ActiveRecord::Base.send :include, Axlsx::Ar
