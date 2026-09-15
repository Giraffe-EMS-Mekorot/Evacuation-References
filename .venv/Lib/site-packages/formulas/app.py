#!/usr/bin/env python
# -*- coding: UTF-8 -*-
#
# Copyright 2016-2026 European Commission (JRC);
# Licensed under the EUPL (the 'Licence');
# You may not use this work except in compliance with the Licence.
# You may obtain a copy of the Licence at: http://ec.europa.eu/idabc/eupl
"""
It provides Application Factory.
"""
import json
import os

import schedula as sh
from flask import Blueprint, Flask, current_app, jsonify, render_template, request

from .cli import (
    CliError,
    _load_model,
    _build_input_payload,
    _collect_missing_messages,
    _extract_value,
    _parse_render_entry,
    _render_output,
)


def _env_json(name, default):
    raw = os.environ.get(name)
    if not raw:
        return default
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise CliError('Invalid runtime configuration `%s`.' % name) from exc


def _calculate_from_payload(model, payload):
    payload = payload or {}
    if not isinstance(payload, dict):
        raise CliError('JSON request body must be an object.')
    raw_inputs = payload.get('inputs', {})
    if not isinstance(raw_inputs, dict):
        raise CliError('`inputs` must be an object.')
    raw_outputs = payload.get('outputs', [])
    if raw_outputs is None:
        raw_outputs = []
    if not isinstance(raw_outputs, list):
        raise CliError('`outputs` must be a list.')
    raw_renders = payload.get('renders', [])
    if raw_renders is None:
        raw_renders = []
    if not isinstance(raw_renders, list):
        raise CliError('`renders` must be a list.')

    render_specs = [_parse_render_entry(item) for item in raw_renders]
    compute_refs = list(dict.fromkeys(
        [str(item) for item in raw_outputs] + [spec.ref for spec in render_specs]
    ))
    warnings = _collect_missing_messages(model, raw_inputs, compute_refs, render_specs)
    normalized_inputs = _build_input_payload(raw_inputs)
    solution = model.calculate(inputs=normalized_inputs, outputs=compute_refs or None)
    rendered = _render_output(solution, render_specs, compute_refs)
    return {
        'inputs': {
            key: _extract_value(value)
            for key, value in normalized_inputs.items()
        },
        'outputs': rendered,
        'warnings': warnings,
    }


def create_api_blueprint():
    bp = Blueprint('api', __name__, url_prefix='/api')

    @bp.get('/health')
    def health():
        return jsonify({'status': 'ok'})

    @bp.get('/model')
    def model_info():
        return jsonify(current_app.config['FORMULAS_INFO'])

    @bp.post('/calculate')
    def calculate_api():
        try:
            result = _calculate_from_payload(
                current_app.config['FORMULAS_MODEL'], request.get_json(silent=True)
            )
        except CliError as exc:
            return jsonify({'error': exc.format_message()}), 400
        return jsonify(result)

    return bp


def create_gui_blueprint():
    bp = Blueprint('gui', __name__, template_folder='gui/templates')

    @bp.get('/')
    def index():
        return render_template(
            'index.html',
            info=current_app.config['FORMULAS_INFO'],
            result=None,
            error=None,
        )

    return bp


def create_app(files=None, circular=None, gui=True):
    if files is None:
        files = tuple(_env_json('FORMULAS_SERVE_FILES', []))
    if circular is None:
        circular = bool(_env_json('FORMULAS_SERVE_CIRCULAR', False))
    if not files:
        raise CliError('At least one input file is required.')
    model = _load_model(files, circular=circular)
    app = Flask(__name__)
    app.config['FORMULAS_MODEL'] = model
    app.config['FORMULAS_INFO'] = {
        'files': list(files),
        'books': sorted(model.books),
        'nodes': len([
            node for node in model.dsp.data_nodes if not isinstance(node, sh.Token)
        ]),
    }
    app.register_blueprint(create_api_blueprint())
    if gui:
        app.register_blueprint(create_gui_blueprint())
    return app
