/**
 * mock-api.src.js — Dataverse metadata API mock for the PCF test harness.
 *
 * patch.js copies this file to node_modules/pcf-start/lib/mock-api.js after
 * every `npm install`. It patches window.fetch to intercept /api/data/v9.2/*
 * requests and return realistic mock data for all 5 metadata types so the
 * control can be tested in the pcf-start harness without a live Dataverse
 * connection.
 *
 * To update mock data: edit THIS file, then run `node patch.js`.
 */
(function () {
  'use strict';

  // ── Helpers ─────────────────────────────────────────────────────────────────

  /** Produces a deterministic-looking GUID from a small integer seed. */
  function guid(n) {
    var h = n.toString(16).padStart(8, '0');
    return h + '-' + h.slice(0, 4) + '-4' + h.slice(1, 4) + '-a' + h.slice(2, 5) + '-' + h.padStart(12, '0');
  }

  function attr(logicalName, label, type, seed) {
    return {
      LogicalName: logicalName,
      DisplayName: { UserLocalizedLabel: { Label: label } },
      MetadataId: guid(seed),
      AttributeType: type,
    };
  }

  function entity(logicalName, label, seed) {
    return {
      LogicalName: logicalName,
      DisplayName: { UserLocalizedLabel: { Label: label } },
      MetadataId: guid(seed),
    };
  }

  function rel(referencingEntity, seed) {
    return { ReferencingEntity: referencingEntity, MetadataId: guid(seed) };
  }

  function view(name, id) {
    return { name: name, savedqueryid: id };
  }

  function bpf(name, id) {
    return { name: name, workflowid: id };
  }

  // ── Mock data ────────────────────────────────────────────────────────────────

  var ENTITIES = [
    entity('account',          'Account',           1),
    entity('contact',          'Contact',           2),
    entity('lead',             'Lead',              3),
    entity('opportunity',      'Opportunity',       4),
    entity('incident',         'Case',              5),
    entity('task',             'Task',              6),
    entity('email',            'Email',             7),
    entity('phonecall',        'Phone Call',        8),
    entity('appointment',      'Appointment',       9),
    entity('systemuser',       'User',              10),
    entity('team',             'Team',              11),
    entity('businessunit',     'Business Unit',     12),
    entity('quote',            'Quote',             13),
    entity('salesorder',       'Order',             14),
    entity('invoice',          'Invoice',           15),
    entity('product',          'Product',           16),
    entity('campaign',         'Campaign',          17),
    entity('list',             'Marketing List',    18),
    entity('knowledgearticle', 'Knowledge Article', 19),
    entity('activitypointer',  'Activity',          20),
  ];

  // Account attributes
  var ACCOUNT_ATTRS = [
    attr('accountid',                    'Account',                   'Uniqueidentifier', 101),
    attr('name',                         'Account Name',              'String',           102),
    attr('emailaddress1',                'Email',                     'String',           103),
    attr('telephone1',                   'Business Phone',            'String',           104),
    attr('telephone2',                   'Other Phone',               'String',           105),
    attr('fax',                          'Fax',                       'String',           106),
    attr('websiteurl',                   'Website',                   'String',           107),
    attr('revenue',                      'Annual Revenue',            'Money',            108),
    attr('numberofemployees',            'No. of Employees',          'Integer',          109),
    attr('description',                  'Description',               'Memo',             110),
    attr('parentaccountid',              'Parent Account',            'Lookup',           111),
    attr('primarycontactid',             'Primary Contact',           'Lookup',           112),
    attr('ownerid',                      'Owner',                     'Owner',            113),
    attr('statecode',                    'Status',                    'State',            114),
    attr('statuscode',                   'Status Reason',             'Status',           115),
    attr('createdon',                    'Created On',                'DateTime',         116),
    attr('modifiedon',                   'Modified On',               'DateTime',         117),
    attr('createdby',                    'Created By',                'Lookup',           118),
    attr('modifiedby',                   'Modified By',               'Lookup',           119),
    attr('accountcategorycode',          'Category',                  'Picklist',         120),
    attr('industrycode',                 'Industry',                  'Picklist',         121),
    attr('sic',                          'SIC Code',                  'String',           122),
    attr('address1_line1',               'Street 1',                  'String',           123),
    attr('address1_city',                'City',                      'String',           124),
    attr('address1_stateorprovince',     'State/Province',            'String',           125),
    attr('address1_postalcode',          'ZIP/Postal Code',           'String',           126),
    attr('address1_country',             'Country/Region',            'String',           127),
    attr('creditlimit',                  'Credit Limit',              'Money',            128),
    attr('creditonhold',                 'Credit Hold',               'Boolean',          129),
    attr('preferredcontactmethodcode',   'Preferred Contact Method',  'Picklist',         130),
    attr('customersizecode',             'Customer Size',             'Picklist',         131),
    attr('originatingleadid',            'Originating Lead',          'Lookup',           132),
    attr('masterid',                     'Master ID',                 'Lookup',           133),
  ];

  // Contact attributes
  var CONTACT_ATTRS = [
    attr('contactid',        'Contact',          'Uniqueidentifier', 201),
    attr('firstname',        'First Name',       'String',           202),
    attr('lastname',         'Last Name',        'String',           203),
    attr('fullname',         'Full Name',        'String',           204),
    attr('emailaddress1',    'Email',            'String',           205),
    attr('telephone1',       'Business Phone',   'String',           206),
    attr('mobilephone',      'Mobile Phone',     'String',           207),
    attr('jobtitle',         'Job Title',        'String',           208),
    attr('department',       'Department',       'String',           209),
    attr('parentcustomerid', 'Company Name',     'Customer',         210),
    attr('parentaccountid',  'Account',          'Lookup',           211),
    attr('ownerid',          'Owner',            'Owner',            212),
    attr('statecode',        'Status',           'State',            213),
    attr('statuscode',       'Status Reason',    'Status',           214),
    attr('createdon',        'Created On',       'DateTime',         215),
    attr('modifiedon',       'Modified On',      'DateTime',         216),
    attr('birthdate',        'Birthday',         'DateTime',         217),
    attr('gendercode',       'Gender',           'Picklist',         218),
    attr('address1_city',    'City',             'String',           219),
    attr('address1_country', 'Country/Region',   'String',           220),
  ];

  // Generic attributes (fallback for unlisted entities)
  var GENERIC_ATTRS = [
    attr('id',          'ID',            'Uniqueidentifier', 901),
    attr('name',        'Name',          'String',           902),
    attr('ownerid',     'Owner',         'Owner',            903),
    attr('statecode',   'Status',        'State',            904),
    attr('statuscode',  'Status Reason', 'Status',           905),
    attr('createdon',   'Created On',    'DateTime',         906),
    attr('modifiedon',  'Modified On',   'DateTime',         907),
  ];

  function lookupOnly(attrs) {
    return attrs.filter(function (a) {
      return a.AttributeType === 'Lookup' || a.AttributeType === 'Customer' || a.AttributeType === 'Owner';
    });
  }

  // 1:N relationships for 'account'
  var ACCOUNT_1N = [
    rel('contact',     301),
    rel('opportunity', 302),
    rel('incident',    303),
    rel('task',        304),
    rel('email',       305),
    rel('phonecall',   306),
    rel('appointment', 307),
    rel('lead',        308),
    rel('quote',       309),
    rel('salesorder',  310),
    rel('invoice',     311),
  ];

  // System views
  var ACCOUNT_VIEWS = [
    view('Active Accounts',                                   '00000000-0000-0000-00aa-000010001001'),
    view('Inactive Accounts',                                 '00000000-0000-0000-00aa-000010001002'),
    view('My Active Accounts',                                '00000000-0000-0000-00aa-000010001003'),
    view('All Accounts',                                      '00000000-0000-0000-00aa-000010001004'),
    view('Accounts I Follow',                                 '00000000-0000-0000-00aa-000010001005'),
    view('Accounts: No Campaign Activities in Last 3 Months','00000000-0000-0000-00aa-000010001006'),
    view('Accounts: Responded to Campaigns in Last 6 Months','00000000-0000-0000-00aa-000010001007'),
    view('Accounts with Overdue Activities',                  '00000000-0000-0000-00aa-000010001008'),
  ];

  var CONTACT_VIEWS = [
    view('Active Contacts',    '00000000-0000-0000-00bb-000010001001'),
    view('Inactive Contacts',  '00000000-0000-0000-00bb-000010001002'),
    view('My Active Contacts', '00000000-0000-0000-00bb-000010001003'),
    view('All Contacts',       '00000000-0000-0000-00bb-000010001004'),
    view('Contacts I Follow',  '00000000-0000-0000-00bb-000010001005'),
  ];

  var GENERIC_VIEWS = [
    view('Active Records',   '00000000-0000-0000-00cc-000010001001'),
    view('Inactive Records', '00000000-0000-0000-00cc-000010001002'),
    view('My Records',       '00000000-0000-0000-00cc-000010001003'),
    view('All Records',      '00000000-0000-0000-00cc-000010001004'),
  ];

  // Business process flows
  var BPFS = {
    opportunity: [
      bpf('Opportunity Sales Process',         '11111111-1111-1111-1111-111111111111'),
      bpf('Lead to Opportunity Sales Process', '22222222-2222-2222-2222-222222222222'),
    ],
    lead: [
      bpf('Lead to Opportunity Sales Process', '22222222-2222-2222-2222-222222222222'),
      bpf('Lead Qualification Process',        '33333333-3333-3333-3333-333333333333'),
    ],
    incident: [
      bpf('Phone to Case Process',             '44444444-4444-4444-4444-444444444444'),
      bpf('Email to Case Process',             '55555555-5555-5555-5555-555555555555'),
    ],
    contact: [
      bpf('Contact to Case Process',           '66666666-6666-6666-6666-666666666666'),
    ],
  };

  // ── URL matching ─────────────────────────────────────────────────────────────

  function extractEntity(url) {
    var m = url.match(/LogicalName='([^']+)'/i)
           || url.match(/returnedtypecode eq '([^']+)'/i)
           || url.match(/primaryentity eq '([^']+)'/i);
    return m ? m[1].toLowerCase() : null;
  }

  function respond(data) {
    return Promise.resolve(new Response(JSON.stringify(data), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    }));
  }

  function getMock(url) {
    if (url.indexOf('/api/data/') === -1) return null;

    var ent = extractEntity(url);

    // 1. Entity list  (no LogicalName= in URL)
    if (/\/EntityDefinitions\?/.test(url) && !/LogicalName=/.test(url)) {
      return respond({ value: ENTITIES });
    }

    // 2. 1:N relationships  →  Entity type with filterEntityFieldByEntitiesAssociatedTo
    if (/\/OneToManyRelationships/.test(url)) {
      return respond({ value: ACCOUNT_1N });
    }

    // 3. Attribute list  →  Attributes type
    if (/\/Attributes\?/.test(url)) {
      var aList = ent === 'contact' ? CONTACT_ATTRS : ent === 'account' ? ACCOUNT_ATTRS : GENERIC_ATTRS;
      return respond({ value: aList });
    }

    // 4. Lookup/Customer attrs only  →  Lookup type
    if (/\$expand=Attributes/.test(url)) {
      var base = ent === 'contact' ? CONTACT_ATTRS : ent === 'account' ? ACCOUNT_ATTRS : GENERIC_ATTRS;
      return respond({ Attributes: lookupOnly(base) });
    }

    // 5. Saved queries  →  SystemViews type
    if (/\/savedqueries\?/.test(url)) {
      var vList = ent === 'contact' ? CONTACT_VIEWS : ent === 'account' ? ACCOUNT_VIEWS : GENERIC_VIEWS;
      return respond({ value: vList });
    }

    // 6. Workflows  →  BusinessProcessFlows type
    if (/\/workflows\?/.test(url)) {
      return respond({ value: (ent && BPFS[ent]) ? BPFS[ent] : [] });
    }

    return null;
  }

  // ── Patch window.fetch ───────────────────────────────────────────────────────

  var _origFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {
    var url = typeof input === 'string' ? input
            : (input && typeof input.url === 'string') ? input.url
            : String(input);
    var mock = getMock(url);
    if (mock) {
      var label = url.split('/api/data/')[1] || url;
      console.log('[Mock API] \u2192', label.split('?')[0]);
      return mock;
    }
    return _origFetch(input, init);
  };

  console.log('[Mock API] installed \u2014 /api/data/* calls are mocked in this harness');
}());
