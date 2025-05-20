# about: This file contains the function used by FetchData class to format data

class FetchUtils:

    def add_final_environment(self, L):
        final_env = {"Production": "Prod"}
        i = L[0].index("Environment")
        L[0].insert(i + 1, "Final Environment")
        for l in L[1:]:
            try:
                l.insert(i + 1, final_env[l[i].strip()])
            except KeyError as ke:
                l.insert(i + 1, "Non-Prod")
            except Exception as e:
                l.insert(i + 1, "Non-Prod")

        return L

    def add_final_state(self, L):
        final_state = {"Resolved": "Closed", "Closed": "Closed"}
        i = L[0].index("State")
        L[0].insert(i + 1, "Final State")
        for l in L[1:]:
            try:
                l.insert(i + 1, final_state[l[i].strip()])
            except KeyError as ke:
                l.insert(i + 1, "Open")
            except Exception as e:
                print(e)

        return L

    def fix_category_subcategory_blanks(self, L):
        i = L[0].index("Category")
        j = L[0].index("Subcategory")
        for l in L[1:]:
            try:
                if l[i] == "" or l[i] is None:
                    l[i] = "Other"
                    l[j] = "Other"
                elif l[j] == "" or l[j] is None:
                    l[j] = "Other"

            except Exception as e:
                print(e)

        return L

    def add_bep(self, L):
        i = L[0].index("Business elapsed time")
        L[0].insert(i + 1, "BET in hrs")
        for l in L[1:]:
            try:
                l.insert(i + 1, l[i] / 3600)
            except KeyError:
                l.insert(i + 1, 'NA')
            except Exception as e:
                print(e)

        return L

    def add_final_assignment_group(self, L):
        index_assign_group = L[0].index('Assignment group')
        index_self_managed = L[0].index('Launchpad Project')
        index_project_name = L[0].index('Project Name')
        num_index = L[0].index('Number')
        L[0].insert(index_assign_group + 1, "Final Assignment Group")  # adding header for row

        for l in L[1:]:
            project_name = l[index_project_name].strip()
            assign_group_name = l[index_assign_group].strip()
            self_managed = l[index_self_managed]
            final_assign_group = 'NA'

            if project_name.__contains__("Ninja") or assign_group_name.__contains__("Ninja"):
                final_assign_group = "DevOps Ninja"
            elif project_name.__contains__("Network Rail Operations") or assign_group_name.__contains__(
                    "Network Rail Operations"):
                final_assign_group = "CMS UK"
            elif assign_group_name.startswith("LP") or assign_group_name.__contains__(
                    "Launchpad") or self_managed == True:
                final_assign_group = "CMS Launchpad"
            elif assign_group_name == 'AIM LS CogX AMS Team' or assign_group_name.__contains__("Analytics DevOps"):
                final_assign_group = "CMS X-POD"
            elif assign_group_name.__contains__("OpenCloud"):
                final_assign_group = "CMS OpenCloud"
            elif assign_group_name.startswith("CMS UK") or assign_group_name in ['Network Rail Operations',
                                                                                 'Model Trade Platform Operations']:
                final_assign_group = "CMS UK"
            elif assign_group_name.__contains__(
                    'CMS X-POD') or assign_group_name == 'Spirits Aero - AMS Support' or assign_group_name.__contains__(
                    "Cross Industry-Internal DEVOPS"):
                final_assign_group = 'CMS X-POD'
            elif assign_group_name.startswith("Daiichi"):
                final_assign_group = "Daiichi"
            elif assign_group_name in ["DCloud Analytics", "GPS SLG DevOps"]:
                final_assign_group = "GPS SLG DevOps"
            elif assign_group_name in ["DeloitteSAPPod - GCP Support", 'Gallo SAP Basis', 'GFS SAP Basis', 'SAP DEVOPS',
                                       'Stericycle SAP Basis Team', 'SAP Splunk TOSCA']:
                final_assign_group = 'SAP DEVOPS'
            elif assign_group_name.startswith('Security Services'):
                final_assign_group = 'Security Services'
            elif assign_group_name in ['OCI Infra DevOps', 'McD J AWS Infra DevOps']:
                final_assign_group = 'OCI Infra DevOps'
            elif assign_group_name == 'E-Commerce DevOps':
                final_assign_group = 'E-Commerce DevOps'
            elif assign_group_name == 'OCI EBS DBA DevOps':
                final_assign_group = 'OCI EBS DBA DevOps'
            elif assign_group_name == 'CMS US Support DevOps':
                final_assign_group = 'CMS US Support DevOps'
            elif assign_group_name == 'MSIX Operations':
                final_assign_group = 'MSIX Operations'
            elif assign_group_name == 'CMS Compliance Team':
                final_assign_group = 'CMS Compliance Team'
            elif assign_group_name == 'Global Cloud Operations':
                final_assign_group = 'Global Cloud Operations'
            else:
                final_assign_group = assign_group_name

            l.insert(index_assign_group + 1, final_assign_group)  # inserting final assignment group value

        return L
